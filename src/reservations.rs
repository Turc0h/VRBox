//! Application service for grouped, multi-station bookings.
//! A group either reserves every requested station or leaves the database untouched.

use chrono::NaiveDate;
use rusqlite::{Connection, params};
use std::collections::BTreeSet;

pub struct NewReservationGroup<'a> {
    pub day: &'a str,
    pub customer: &'a str,
    pub start_hour: i32,
    pub promo_name: &'a str,
    /// Each station carries its net one-hour price and allocated discount in cents.
    pub stations: &'a [(&'a str, i64, i64)],
}

pub struct UpdateReservationGroup<'a> {
    pub group_id: i64,
    pub day: &'a str,
    pub customer: &'a str,
    pub start_hour: i32,
    pub promo_name: &'a str,
    pub stations: &'a [(&'a str, i64, i64)],
}

pub struct UpdateSingleReservation<'a> {
    pub reservation_id: i64,
    pub day: &'a str,
    pub customer: &'a str,
    pub start_hour: i32,
    pub station: &'a str,
    pub total_cents: i64,
    pub discount_cents: i64,
}

#[derive(Clone, Copy)]
pub enum PaymentAction {
    Deposit,
    Balance,
    Partial,
}

fn validate_group(
    day: &str,
    customer: &str,
    start_hour: i32,
    stations: &[(&str, i64, i64)],
) -> Result<i64, String> {
    if customer.trim().is_empty() {
        return Err("Ingresá el nombre del cliente.".into());
    }
    if NaiveDate::parse_from_str(day, "%Y-%m-%d").is_err() {
        return Err("La fecha debe tener formato AAAA-MM-DD.".into());
    }
    if !(16..=23).contains(&start_hour) {
        return Err("El horario debe estar entre las 16:00 y las 23:00.".into());
    }
    if stations.is_empty() {
        return Err("Seleccioná al menos una estación.".into());
    }
    let mut unique = BTreeSet::new();
    for (resource, price, discount) in stations {
        if resource.trim().is_empty() || *price < 0 || *discount < 0 {
            return Err("La estación o sus importes no son válidos.".into());
        }
        if !unique.insert(*resource) {
            return Err("No se puede repetir una estación dentro de la reserva.".into());
        }
    }
    stations
        .iter()
        .try_fold(0_i64, |sum, (_, price, _)| sum.checked_add(*price))
        .ok_or("El total de la reserva supera el máximo admitido.".into())
}

fn net_paid(
    db: &rusqlite::Connection,
    group_id: Option<i64>,
    reservation_id: Option<i64>,
) -> Result<i64, String> {
    let value: i64 = if let Some(id) = group_id {
        db.query_row("SELECT COALESCE(SUM(CASE WHEN kind='devolucion' THEN -amount_cents ELSE amount_cents END),0) FROM reservation_payments WHERE group_id=?1", [id], |r| r.get(0))
    } else if let Some(id) = reservation_id {
        db.query_row("SELECT COALESCE(SUM(CASE WHEN kind='devolucion' THEN -amount_cents ELSE amount_cents END),0) FROM reservation_payments WHERE reservation_id=?1", [id], |r| r.get(0))
    } else { return Err("La reserva no es válida.".into()); }.map_err(|e| format!("No se pudo leer los pagos: {e}"))?;
    Ok(value)
}

pub fn record_payment(
    db: &mut Connection,
    group_id: Option<i64>,
    reservation_id: Option<i64>,
    action: PaymentAction,
    requested_cents: i64,
    method: &str,
    service_date: &str,
) -> Result<i64, String> {
    if group_id.is_some() == reservation_id.is_some() {
        return Err("La reserva no es válida.".into());
    }
    if !["efectivo", "transferencia"].contains(&method) {
        return Err("El medio de pago no es válido.".into());
    }
    if NaiveDate::parse_from_str(service_date, "%Y-%m-%d").is_err() {
        return Err("La fecha del pago no es válida.".into());
    }
    let tx = db
        .transaction_with_behavior(rusqlite::TransactionBehavior::Immediate)
        .map_err(|e| format!("No se pudo iniciar el pago: {e}"))?;
    let (total, active): (i64, i64) = if let Some(id) = group_id {
        tx.query_row("SELECT g.total_cents,(SELECT COUNT(*) FROM reservations r WHERE r.group_id=g.id AND r.status IN ('reservada','completada')) FROM reservation_groups g WHERE g.id=?1", [id], |r| Ok((r.get(0)?,r.get(1)?))).map_err(|_| "No se encontró la reserva.".to_string())?
    } else {
        tx.query_row("SELECT total_cents,CASE WHEN status IN ('reservada','completada') THEN 1 ELSE 0 END FROM reservations WHERE id=?1", [reservation_id.unwrap()], |r| Ok((r.get(0)?,r.get(1)?))).map_err(|_| "No se encontró la reserva.".to_string())?
    };
    if active == 0 {
        return Err("La reserva está cancelada y no admite nuevos pagos.".into());
    }
    let paid = net_paid(&tx, group_id, reservation_id)?;
    let balance = total.saturating_sub(paid);
    let amount = match action {
        PaymentAction::Deposit => {
            let target = (total + 1) / 2;
            if paid >= target {
                return Err(
                    "La seña del 50% ya está cubierta. Registrá el saldo o un pago parcial.".into(),
                );
            }
            target - paid
        }
        PaymentAction::Balance => {
            if balance <= 0 {
                return Err("La reserva ya está pagada.".into());
            }
            balance
        }
        PaymentAction::Partial => {
            if requested_cents <= 0 || requested_cents > balance {
                return Err(format!(
                    "El pago parcial debe ser mayor que cero y no superar el saldo de {}.",
                    balance
                ));
            }
            requested_cents
        }
    };
    let kind = match action {
        PaymentAction::Deposit => "sena",
        PaymentAction::Balance => "saldo",
        PaymentAction::Partial => "parcial",
    };
    let target = group_id
        .map(|id| format!("grupo #{id}"))
        .unwrap_or_else(|| format!("reserva #{}", reservation_id.unwrap()));
    tx.execute("INSERT INTO sales(category,item,quantity,unit_cents,total_cents,payment,service_date) VALUES('reserva',?1,1,?2,?2,?3,?4)", params![format!("{kind} · {target}"), amount, method, service_date]).map_err(|e| format!("No se pudo registrar el ingreso: {e}"))?;
    let sale_id = tx.last_insert_rowid();
    tx.execute("INSERT INTO reservation_payments(group_id,reservation_id,amount_cents,sale_id,kind,payment,service_date) VALUES(?1,?2,?3,?4,?5,?6,?7)", params![group_id,reservation_id,amount,sale_id,kind,method,service_date]).map_err(|e| format!("No se pudo vincular el pago con la reserva: {e}"))?;
    tx.commit()
        .map_err(|e| format!("No se pudo confirmar el pago: {e}"))?;
    Ok(amount)
}

pub fn cancel_group(
    db: &mut Connection,
    group_id: Option<i64>,
    reservation_id: Option<i64>,
    method: &str,
    service_date: &str,
) -> Result<i64, String> {
    if group_id.is_some() == reservation_id.is_some() {
        return Err("La reserva no es válida.".into());
    }
    let tx = db
        .transaction_with_behavior(rusqlite::TransactionBehavior::Immediate)
        .map_err(|e| format!("No se pudo iniciar la cancelación: {e}"))?;
    let blocked: i64 = if let Some(id) = group_id {
        tx.query_row(
            "SELECT COUNT(*) FROM reservations WHERE group_id=?1 AND status!='reservada'",
            [id],
            |r| r.get(0),
        )
    } else {
        tx.query_row(
            "SELECT COUNT(*) FROM reservations WHERE id=?1 AND status!='reservada'",
            [reservation_id.unwrap()],
            |r| r.get(0),
        )
    }
    .map_err(|e| format!("No se pudo comprobar el estado: {e}"))?;
    let exists: i64 = if let Some(id) = group_id {
        tx.query_row(
            "SELECT COUNT(*) FROM reservations WHERE group_id=?1",
            [id],
            |r| r.get(0),
        )
    } else {
        tx.query_row(
            "SELECT COUNT(*) FROM reservations WHERE id=?1",
            [reservation_id.unwrap()],
            |r| r.get(0),
        )
    }
    .map_err(|e| format!("No se pudo comprobar la reserva: {e}"))?;
    if exists == 0 {
        return Err("No se encontró la reserva.".into());
    }
    if blocked > 0 {
        return Err("Solo se pueden cancelar reservas que todavía no comenzaron.".into());
    }
    let paid = net_paid(&tx, group_id, reservation_id)?;
    if paid > 0 {
        if !["efectivo", "transferencia"].contains(&method) {
            return Err("Elegí el medio usado para devolver el dinero.".into());
        }
        tx.execute("INSERT INTO sales(category,item,quantity,unit_cents,total_cents,payment,service_date) VALUES('devolucion',?1,1,?2,?3,?4,?5)", params![format!("Devolución reserva {}", group_id.map(|id| format!("#{id}")).unwrap_or_else(|| format!("#{}", reservation_id.unwrap()))), paid, -paid, method, service_date]).map_err(|e| format!("No se pudo registrar la devolución: {e}"))?;
        let sale_id = tx.last_insert_rowid();
        tx.execute("INSERT INTO reservation_payments(group_id,reservation_id,amount_cents,sale_id,kind,payment,service_date) VALUES(?1,?2,?3,?4,'devolucion',?5,?6)", params![group_id,reservation_id,paid,sale_id,method,service_date]).map_err(|e| format!("No se pudo guardar la devolución: {e}"))?;
    }
    let changed = if let Some(id) = group_id {
        tx.execute(
            "UPDATE reservations SET status='cancelada' WHERE group_id=?1 AND status='reservada'",
            [id],
        )
    } else {
        tx.execute(
            "UPDATE reservations SET status='cancelada' WHERE id=?1 AND status='reservada'",
            [reservation_id.unwrap()],
        )
    }
    .map_err(|e| format!("No se pudo cancelar la reserva: {e}"))?;
    if changed == 0 {
        return Err("Solo se pueden cancelar reservas aún no iniciadas.".into());
    }
    if let Some(id) = group_id {
        tx.execute(
            "INSERT INTO reservation_audit(group_id,action,details) VALUES(?1,'cancelada',?2)",
            params![
                id,
                format!("Cancelada; devolución registrada: {paid} centavos.")
            ],
        )
        .map_err(|e| format!("No se pudo guardar el historial: {e}"))?;
    } else if let Some(id) = reservation_id {
        tx.execute(
            "INSERT INTO reservation_audit(reservation_id,action,details) VALUES(?1,'cancelada',?2)",
            params![id, format!("Cancelada; devolución registrada: {paid} centavos.")],
        )
        .map_err(|e| format!("No se pudo guardar el historial: {e}"))?;
    }
    tx.commit()
        .map_err(|e| format!("No se pudo confirmar la cancelación: {e}"))?;
    Ok(paid)
}

pub fn create_group(db: &mut Connection, booking: NewReservationGroup<'_>) -> Result<i64, String> {
    if booking.customer.trim().is_empty() {
        return Err("Ingresá el nombre del cliente.".into());
    }
    if NaiveDate::parse_from_str(booking.day, "%Y-%m-%d").is_err() {
        return Err("La fecha debe tener formato AAAA-MM-DD.".into());
    }
    if !(16..=23).contains(&booking.start_hour) {
        return Err("El horario debe estar entre las 16:00 y las 23:00.".into());
    }
    if booking.stations.is_empty() {
        return Err("Seleccioná al menos una estación.".into());
    }
    let mut unique_stations = BTreeSet::new();
    for (resource, price, discount) in booking.stations {
        if resource.trim().is_empty() || *price < 0 || *discount < 0 {
            return Err("La estación o su tarifa no son válidas.".into());
        }
        if !unique_stations.insert(*resource) {
            return Err("No se puede repetir una estación dentro de la misma reserva.".into());
        }
    }

    let quote = booking
        .stations
        .iter()
        .try_fold(0_i64, |sum, (_, price, _)| sum.checked_add(*price))
        .ok_or("El total de la reserva supera el máximo admitido.")?;
    let discount = booking
        .stations
        .iter()
        .try_fold(0_i64, |sum, (_, _, discount)| sum.checked_add(*discount))
        .ok_or("El descuento supera el máximo admitido.")?;
    let tx = db
        .transaction_with_behavior(rusqlite::TransactionBehavior::Immediate)
        .map_err(|e| format!("No se pudo iniciar la reserva: {e}"))?;

    for (resource, _, _) in booking.stations {
        let collision: i64 = tx
            .query_row(
                "SELECT COUNT(*) FROM reservations
                 WHERE service_date=?1 AND resource=?2 AND status!='cancelada'
                   AND start_hour < ?3 AND start_hour + duration > ?4",
                params![
                    booking.day,
                    resource,
                    booking.start_hour + 1,
                    booking.start_hour
                ],
                |row| row.get(0),
            )
            .map_err(|e| format!("No se pudo verificar la disponibilidad: {e}"))?;
        if collision > 0 {
            return Err(format!(
                "{} ya está ocupada a las {:02}:00. No se guardó ninguna estación.",
                resource, booking.start_hour
            ));
        }
    }

    tx.execute(
        "INSERT INTO reservation_groups(service_date,customer,start_hour,station_count,total_cents,promo_name,discount_cents)
         VALUES(?1,?2,?3,?4,?5,?6,?7)",
        params![
            booking.day,
            booking.customer.trim(),
            booking.start_hour,
            booking.stations.len() as i64,
            quote,
            booking.promo_name,
            discount
        ],
    )
    .map_err(|e| format!("No se pudo guardar la reserva: {e}"))?;
    let group_id = tx.last_insert_rowid();

    for (resource, price, line_discount) in booking.stations {
        tx.execute(
            "INSERT INTO reservations(
                service_date,resource,customer,start_hour,duration,games,total_cents,
                discount_cents,note,status,group_id
             ) VALUES(?1,?2,?3,?4,1,'',?5,?6,'','reservada',?7)",
            params![
                booking.day,
                resource,
                booking.customer.trim(),
                booking.start_hour,
                price,
                line_discount,
                group_id
            ],
        )
        .map_err(|e| format!("No se pudo completar el grupo de estaciones: {e}"))?;
    }

    tx.commit()
        .map_err(|e| format!("No se pudo confirmar la reserva: {e}"))?;
    Ok(group_id)
}

pub fn update_group(
    db: &mut Connection,
    booking: UpdateReservationGroup<'_>,
) -> Result<(), String> {
    let quote = validate_group(
        booking.day,
        booking.customer,
        booking.start_hour,
        booking.stations,
    )?;
    let discount = booking
        .stations
        .iter()
        .try_fold(0_i64, |sum, (_, _, d)| sum.checked_add(*d))
        .ok_or("El descuento supera el máximo admitido.")?;
    let tx = db
        .transaction_with_behavior(rusqlite::TransactionBehavior::Immediate)
        .map_err(|e| format!("No se pudo iniciar la modificación: {e}"))?;
    let (old_day, old_customer, old_hour, old_total, status_count): (String, String, i32, i64, i64) = tx.query_row(
        "SELECT g.service_date,g.customer,g.start_hour,g.total_cents,(SELECT COUNT(*) FROM reservations r WHERE r.group_id=g.id AND r.status!='reservada') FROM reservation_groups g WHERE g.id=?1",
        [booking.group_id], |r| Ok((r.get(0)?,r.get(1)?,r.get(2)?,r.get(3)?,r.get(4)?))
    ).map_err(|_| "No se encontró el grupo de reserva.".to_string())?;
    if status_count > 0 {
        return Err("Solo se pueden modificar reservas que todavía no comenzaron.".into());
    }
    let paid = net_paid(&tx, Some(booking.group_id), None)?;
    if paid > quote {
        return Err(format!(
            "El nuevo total {} queda por debajo de lo ya pagado {}. Registrá una devolución antes de reducir el precio.",
            quote, paid
        ));
    }
    let old: Vec<(i64, String)> = {
        let mut stmt = tx
            .prepare("SELECT id,resource FROM reservations WHERE group_id=?1 ORDER BY id")
            .map_err(|e| e.to_string())?;
        stmt.query_map([booking.group_id], |r| Ok((r.get(0)?, r.get(1)?)))
            .map_err(|e| e.to_string())?
            .collect::<Result<Vec<_>, _>>()
            .map_err(|e| e.to_string())?
    };
    for (resource, _, _) in booking.stations {
        let collision: i64 = tx.query_row(
            "SELECT COUNT(*) FROM reservations WHERE service_date=?1 AND resource=?2 AND status!='cancelada' AND COALESCE(group_id,0)!=?3 AND start_hour < ?4 AND start_hour + duration > ?5",
            params![booking.day, resource, booking.group_id, booking.start_hour+1, booking.start_hour], |r| r.get(0)
        ).map_err(|e| format!("No se pudo verificar la disponibilidad: {e}"))?;
        if collision > 0 {
            return Err(format!(
                "{} ya está ocupada a las {:02}:00. No se modificó la reserva.",
                resource, booking.start_hour
            ));
        }
    }
    tx.execute("UPDATE reservation_groups SET service_date=?1,customer=?2,start_hour=?3,station_count=?4,total_cents=?5,promo_name=?6,discount_cents=?7 WHERE id=?8",
        params![booking.day,booking.customer.trim(),booking.start_hour,booking.stations.len() as i64,quote,booking.promo_name,discount,booking.group_id]
    ).map_err(|e| format!("No se pudo modificar el grupo: {e}"))?;
    let mut used = BTreeSet::new();
    for (resource, price, line_discount) in booking.stations {
        if let Some((id, _)) = old
            .iter()
            .find(|(_, old_resource)| old_resource == resource)
        {
            tx.execute("UPDATE reservations SET service_date=?1,customer=?2,start_hour=?3,total_cents=?4,discount_cents=?5 WHERE id=?6",
                params![booking.day,booking.customer.trim(),booking.start_hour,price,line_discount,id]
            ).map_err(|e| format!("No se pudo actualizar una estación: {e}"))?;
            used.insert(*id);
        } else {
            tx.execute("INSERT INTO reservations(service_date,resource,customer,start_hour,duration,games,total_cents,discount_cents,note,status,group_id) VALUES(?1,?2,?3,?4,1,'',?5,?6,'','reservada',?7)",
                params![booking.day,resource,booking.customer.trim(),booking.start_hour,price,line_discount,booking.group_id]
            ).map_err(|e| format!("No se pudo añadir una estación: {e}"))?;
        }
    }
    for (id, _) in &old {
        if !used.contains(id) {
            tx.execute("DELETE FROM reservations WHERE id=?1", [id])
                .map_err(|e| format!("No se pudo quitar una estación: {e}"))?;
        }
    }
    let old_stations = old
        .iter()
        .map(|(_, name)| name.as_str())
        .collect::<Vec<_>>()
        .join(", ");
    let new_stations = booking
        .stations
        .iter()
        .map(|(name, _, _)| *name)
        .collect::<Vec<_>>()
        .join(", ");
    let details = format!(
        "{} {:02}:00 {} [{}] {} -> {} {:02}:00 {} [{}] {}",
        old_day,
        old_hour,
        old_customer,
        old_stations,
        old_total,
        booking.day,
        booking.start_hour,
        booking.customer.trim(),
        new_stations,
        quote
    );
    tx.execute(
        "INSERT INTO reservation_audit(group_id,action,details) VALUES(?1,'modificada',?2)",
        params![booking.group_id, details],
    )
    .map_err(|e| format!("No se pudo guardar el historial: {e}"))?;
    tx.commit()
        .map_err(|e| format!("No se pudo confirmar la modificación: {e}"))?;
    Ok(())
}

pub fn update_single(
    db: &mut Connection,
    booking: UpdateSingleReservation<'_>,
) -> Result<(), String> {
    if booking.customer.trim().is_empty() {
        return Err("Ingresá el nombre del cliente.".into());
    }
    if NaiveDate::parse_from_str(booking.day, "%Y-%m-%d").is_err() {
        return Err("La fecha debe tener formato AAAA-MM-DD.".into());
    }
    if !(16..=23).contains(&booking.start_hour) {
        return Err("El horario debe estar entre las 16:00 y las 23:00.".into());
    }
    if booking.station.trim().is_empty() || booking.total_cents < 0 || booking.discount_cents < 0 {
        return Err("La estación o sus importes no son válidos.".into());
    }
    let tx = db
        .transaction_with_behavior(rusqlite::TransactionBehavior::Immediate)
        .map_err(|e| format!("No se pudo iniciar la modificación: {e}"))?;
    let (old_day,old_customer,old_resource,old_hour,old_total,status):(String,String,String,i32,i64,String) = tx.query_row(
        "SELECT service_date,customer,resource,start_hour,total_cents,status FROM reservations WHERE id=?1 AND group_id IS NULL",
        [booking.reservation_id], |r| Ok((r.get(0)?,r.get(1)?,r.get(2)?,r.get(3)?,r.get(4)?,r.get(5)?))
    ).map_err(|_| "No se encontró la reserva individual.".to_string())?;
    if status != "reservada" {
        return Err("Solo se pueden modificar reservas que todavía no comenzaron.".into());
    }
    let paid = net_paid(&tx, None, Some(booking.reservation_id))?;
    if paid > booking.total_cents {
        return Err(format!(
            "El nuevo total queda por debajo de lo ya pagado ({}). Registrá una devolución antes de reducir el precio.",
            paid
        ));
    }
    let collision: i64 = tx.query_row(
        "SELECT COUNT(*) FROM reservations WHERE service_date=?1 AND resource=?2 AND status!='cancelada' AND id!=?3 AND start_hour < ?4 AND start_hour + duration > ?5",
        params![booking.day,booking.station,booking.reservation_id,booking.start_hour+1,booking.start_hour], |r| r.get(0)
    ).map_err(|e| format!("No se pudo verificar la disponibilidad: {e}"))?;
    if collision > 0 {
        return Err(format!(
            "{} ya está ocupada a las {:02}:00. No se modificó la reserva.",
            booking.station, booking.start_hour
        ));
    }
    tx.execute("UPDATE reservations SET service_date=?1,resource=?2,customer=?3,start_hour=?4,total_cents=?5,discount_cents=?6 WHERE id=?7",
        params![booking.day,booking.station,booking.customer.trim(),booking.start_hour,booking.total_cents,booking.discount_cents,booking.reservation_id]
    ).map_err(|e| format!("No se pudo actualizar la reserva: {e}"))?;
    let details = format!(
        "{} {:02}:00 {} {} {} -> {} {:02}:00 {} {} {}",
        old_day,
        old_hour,
        old_resource,
        old_customer,
        old_total,
        booking.day,
        booking.start_hour,
        booking.station,
        booking.customer.trim(),
        booking.total_cents
    );
    tx.execute(
        "INSERT INTO reservation_audit(reservation_id,action,details) VALUES(?1,'modificada',?2)",
        params![booking.reservation_id, details],
    )
    .map_err(|e| format!("No se pudo guardar el historial: {e}"))?;
    tx.commit()
        .map_err(|e| format!("No se pudo confirmar la modificación: {e}"))?;
    Ok(())
}
