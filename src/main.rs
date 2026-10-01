use calamine::{Reader, open_workbook_auto};
use chrono::Local;
use eframe::egui::{self, Color32, RichText, Vec2};
use rusqlite::{Connection, params};
use rust_xlsxwriter::{Format, Formula, Workbook};
use std::collections::{BTreeMap, BTreeSet};
use std::{path::PathBuf, time::Duration};
mod reservations;

const RESOURCES: [(&str, &str); 12] = [
    ("Simulador plata 1", "sim_silver"),
    ("Simulador plata 2", "sim_silver"),
    ("Simulador plata 3", "sim_silver"),
    ("Simulador plata 4", "sim_silver"),
    ("Simulador plata 5", "sim_silver"),
    ("Simulador plata 6", "sim_silver"),
    ("Simulador oro 1", "sim_gold"),
    ("Simulador oro 2", "sim_gold"),
    ("PS5 1", "play"),
    ("PS5 2", "play"),
    ("Realidad virtual 1", "vr"),
    ("Realidad virtual 2", "vr"),
];
const GAME_OPTIONS: [&str; 8] = [
    "Assetto Corsa",
    "F1",
    "Forza",
    "Gran Turismo",
    "EA Sports FC",
    "Mortal Kombat",
    "Beat Saber",
    "Otro",
];

fn data_dir() -> PathBuf {
    if let Some(base) = std::env::var_os("LOCALAPPDATA").or_else(|| std::env::var_os("APPDATA")) {
        return PathBuf::from(base).join("VRBox");
    }
    if let Some(home) = std::env::var_os("HOME") {
        return PathBuf::from(home)
            .join(".local")
            .join("share")
            .join("vrbox");
    }
    std::env::current_dir()
        .unwrap_or_else(|_| std::env::temp_dir())
        .join("vrbox-data")
}

fn open_db() -> rusqlite::Result<Connection> {
    let dir = data_dir();
    std::fs::create_dir_all(&dir)
        .map_err(|e| rusqlite::Error::ToSqlConversionFailure(Box::new(e)))?;
    let conn = Connection::open(dir.join("vrbox.sqlite3"))?;
    conn.pragma_update(None, "journal_mode", "WAL")?;
    conn.pragma_update(None, "foreign_keys", "ON")?;
    conn.execute_batch(
        "CREATE TABLE IF NOT EXISTS reservations(
           id INTEGER PRIMARY KEY, service_date TEXT NOT NULL, resource TEXT NOT NULL, customer TEXT NOT NULL,
           start_hour INTEGER NOT NULL, duration INTEGER NOT NULL, games TEXT NOT NULL DEFAULT '',
           total_cents INTEGER NOT NULL, discount_cents INTEGER NOT NULL DEFAULT 0, note TEXT NOT NULL DEFAULT '',
           status TEXT NOT NULL DEFAULT 'reservada', group_id INTEGER, created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
           UNIQUE(service_date, resource, start_hour));
         CREATE TABLE IF NOT EXISTS reservation_groups(
           id INTEGER PRIMARY KEY, service_date TEXT NOT NULL, customer TEXT NOT NULL,
           start_hour INTEGER NOT NULL, station_count INTEGER NOT NULL, total_cents INTEGER NOT NULL,
           promo_name TEXT NOT NULL DEFAULT '', discount_cents INTEGER NOT NULL DEFAULT 0,
           created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP);
         CREATE TABLE IF NOT EXISTS promotions(
           id INTEGER PRIMARY KEY, name TEXT NOT NULL, buy_qty INTEGER NOT NULL CHECK(buy_qty >= 2),
           pay_qty INTEGER NOT NULL CHECK(pay_qty >= 1 AND pay_qty < buy_qty),
           start_hour INTEGER NOT NULL CHECK(start_hour BETWEEN 16 AND 23),
           end_hour INTEGER NOT NULL CHECK(end_hour BETWEEN start_hour AND 23),
           resource_kind TEXT NOT NULL DEFAULT 'all', active INTEGER NOT NULL DEFAULT 1);
         CREATE TABLE IF NOT EXISTS timers(
           id INTEGER PRIMARY KEY, resource TEXT NOT NULL, customer TEXT NOT NULL DEFAULT '',
           started_at INTEGER NOT NULL, ends_at INTEGER NOT NULL, total_secs INTEGER NOT NULL CHECK(total_secs > 0));
         CREATE TABLE IF NOT EXISTS products(
           id INTEGER PRIMARY KEY, name TEXT NOT NULL UNIQUE, price_cents INTEGER NOT NULL CHECK(price_cents >= 0),
           cost_cents INTEGER NOT NULL DEFAULT 0, margin_bps INTEGER NOT NULL DEFAULT 5000,
           stock INTEGER NOT NULL DEFAULT 0 CHECK(stock >= 0), image_path TEXT NOT NULL DEFAULT '', active INTEGER NOT NULL DEFAULT 1);
         CREATE TABLE IF NOT EXISTS sales(
           id INTEGER PRIMARY KEY, sold_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, category TEXT NOT NULL,
           item TEXT NOT NULL, quantity INTEGER NOT NULL, unit_cents INTEGER NOT NULL, total_cents INTEGER NOT NULL,
           payment TEXT NOT NULL CHECK(payment IN ('efectivo','transferencia')), service_date TEXT NOT NULL);
         CREATE TABLE IF NOT EXISTS events(
           id INTEGER PRIMARY KEY, event_date TEXT NOT NULL, kind TEXT NOT NULL, title TEXT NOT NULL,
           customer TEXT NOT NULL DEFAULT '', phone TEXT NOT NULL DEFAULT '', notes TEXT NOT NULL DEFAULT '');
         CREATE TABLE IF NOT EXISTS closings(
           id INTEGER PRIMARY KEY, service_date TEXT NOT NULL UNIQUE, expected_cash INTEGER NOT NULL,
           expected_transfer INTEGER NOT NULL, counted_cash INTEGER NOT NULL, counted_transfer INTEGER NOT NULL,
           closed_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, notes TEXT NOT NULL DEFAULT '');
         CREATE TABLE IF NOT EXISTS reservation_payments(
           id INTEGER PRIMARY KEY, group_id INTEGER REFERENCES reservation_groups(id),
           reservation_id INTEGER REFERENCES reservations(id), amount_cents INTEGER NOT NULL CHECK(amount_cents > 0),
           sale_id INTEGER NOT NULL UNIQUE REFERENCES sales(id),
           kind TEXT NOT NULL CHECK(kind IN ('sena','saldo','parcial','devolucion')),
           payment TEXT NOT NULL CHECK(payment IN ('efectivo','transferencia')),
           service_date TEXT NOT NULL, paid_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
           CHECK((group_id IS NULL) != (reservation_id IS NULL)));
         CREATE INDEX IF NOT EXISTS idx_reservation_payments_group ON reservation_payments(group_id);
         CREATE INDEX IF NOT EXISTS idx_reservation_payments_single ON reservation_payments(reservation_id);
         CREATE TABLE IF NOT EXISTS reservation_audit(
           id INTEGER PRIMARY KEY, group_id INTEGER, reservation_id INTEGER, action TEXT NOT NULL,
           details TEXT NOT NULL, created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
           CHECK((group_id IS NULL) != (reservation_id IS NULL)));
         CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT NOT NULL);")?;
    for (_, kind) in RESOURCES {
        let key = format!("price_{kind}");
        let default = match kind {
            "sim_silver" => "6000",
            "sim_gold" => "8500",
            "play" => "4500",
            _ => "7000",
        };
        conn.execute(
            "INSERT OR IGNORE INTO settings(key,value) VALUES(?1,?2)",
            params![key, default],
        )?;
    }
    let mut columns = conn.prepare("PRAGMA table_info(products)")?;
    let existing = columns
        .query_map([], |r| r.get::<_, String>(1))?
        .filter_map(Result::ok)
        .collect::<Vec<_>>();
    drop(columns);
    if !existing.iter().any(|c| c == "cost_cents") {
        conn.execute(
            "ALTER TABLE products ADD COLUMN cost_cents INTEGER NOT NULL DEFAULT 0",
            [],
        )?;
    }
    if !existing.iter().any(|c| c == "margin_bps") {
        conn.execute(
            "ALTER TABLE products ADD COLUMN margin_bps INTEGER NOT NULL DEFAULT 5000",
            [],
        )?;
    }
    let mut reservation_columns = conn.prepare("PRAGMA table_info(reservations)")?;
    let reservation_existing = reservation_columns
        .query_map([], |r| r.get::<_, String>(1))?
        .filter_map(Result::ok)
        .collect::<Vec<_>>();
    drop(reservation_columns);
    if !reservation_existing.iter().any(|c| c == "group_id") {
        conn.execute("ALTER TABLE reservations ADD COLUMN group_id INTEGER", [])?;
    }
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_reservation_group_id ON reservations(group_id)",
        [],
    )?;
    let mut group_columns = conn.prepare("PRAGMA table_info(reservation_groups)")?;
    let group_existing = group_columns
        .query_map([], |r| r.get::<_, String>(1))?
        .filter_map(Result::ok)
        .collect::<Vec<_>>();
    drop(group_columns);
    if !group_existing.iter().any(|c| c == "promo_name") {
        conn.execute(
            "ALTER TABLE reservation_groups ADD COLUMN promo_name TEXT NOT NULL DEFAULT ''",
            [],
        )?;
    }
    if !group_existing.iter().any(|c| c == "discount_cents") {
        conn.execute(
            "ALTER TABLE reservation_groups ADD COLUMN discount_cents INTEGER NOT NULL DEFAULT 0",
            [],
        )?;
    }
    for index in 1..=6 {
        let _ = conn.execute(
            "UPDATE reservations SET resource=?1 WHERE resource=?2",
            params![
                format!("Simulador plata {index}"),
                format!("Simulador {index}")
            ],
        );
    }
    for index in 1..=2 {
        let _ = conn.execute(
            "UPDATE reservations SET resource=?1 WHERE resource=?2",
            params![
                format!("Simulador oro {index}"),
                format!("Simulador {}", index + 6)
            ],
        );
    }
    conn.execute_batch(
        "WITH ranked_reservations AS (
           SELECT id,group_id,resource,service_date,total_cents,
                  ROW_NUMBER() OVER (PARTITION BY resource,service_date,total_cents ORDER BY start_hour,id) AS rn
           FROM reservations WHERE status='completada'
         ), ranked_sales AS (
           SELECT id,substr(item,1,instr(item,' — ')-1) AS resource,service_date,total_cents,payment,
                  ROW_NUMBER() OVER (PARTITION BY substr(item,1,instr(item,' — ')-1),service_date,total_cents ORDER BY id) AS rn
           FROM sales WHERE category='sesion' AND instr(item,' — ')>0
         )
         INSERT OR IGNORE INTO reservation_payments(group_id,reservation_id,amount_cents,sale_id,kind,payment,service_date)
         SELECT r.group_id,CASE WHEN r.group_id IS NULL THEN r.id ELSE NULL END,s.total_cents,s.id,'saldo',s.payment,s.service_date
         FROM ranked_reservations r JOIN ranked_sales s
           ON s.resource=r.resource AND s.service_date=r.service_date AND s.total_cents=r.total_cents AND s.rn=r.rn;"
    )?;
    conn.execute(
        "INSERT OR IGNORE INTO products(name,price_cents,stock) VALUES('Agua',150000,20)",
        [],
    )?;
    conn.execute(
        "INSERT OR IGNORE INTO products(name,price_cents,stock) VALUES('Gaseosa',250000,12)",
        [],
    )?;
    Ok(conn)
}

fn money(cents: i64) -> String {
    format!("${:.2}", cents as f64 / 100.0)
}
fn remaining_text(seconds: i64) -> String {
    let seconds = seconds.max(0);
    if seconds >= 3600 {
        format!(
            "{:02}:{:02}:{:02}",
            seconds / 3600,
            (seconds / 60) % 60,
            seconds % 60
        )
    } else {
        format!("{:02}:{:02}", seconds / 60, seconds % 60)
    }
}
fn timer_ring(ui: &mut egui::Ui, remaining: i64, total: i64) {
    let (rect, _) = ui.allocate_exact_size(Vec2::splat(116.0), egui::Sense::hover());
    let painter = ui.painter_at(rect);
    let center = rect.center();
    let radius = 43.0;
    painter.circle_stroke(
        center,
        radius,
        egui::Stroke::new(6.0_f32, Color32::from_rgb(43, 59, 77)),
    );
    let fraction = (remaining.max(0) as f32 / total.max(1) as f32).clamp(0.0, 1.0);
    if fraction > 0.0 {
        let segments = (64.0 * fraction).ceil().max(2.0) as usize;
        let points = (0..=segments)
            .map(|i| {
                let angle = -std::f32::consts::FRAC_PI_2
                    + std::f32::consts::TAU * fraction * i as f32 / segments as f32;
                center + Vec2::new(angle.cos(), angle.sin()) * radius
            })
            .collect::<Vec<_>>();
        let color = if remaining == 0 {
            Color32::from_rgb(224, 106, 116)
        } else if fraction < 0.2 {
            Color32::from_rgb(235, 169, 94)
        } else {
            Color32::from_rgb(91, 203, 196)
        };
        painter.add(egui::Shape::line(points, egui::Stroke::new(6.0_f32, color)));
    }
    let text = if remaining == 0 {
        "FINALIZADO".into()
    } else {
        remaining_text(remaining)
    };
    painter.text(
        center,
        egui::Align2::CENTER_CENTER,
        text,
        egui::FontId::proportional(18.0),
        Color32::WHITE,
    );
}
fn pesos_input(pesos: &str) -> i64 {
    pesos
        .replace(',', ".")
        .parse::<f64>()
        .unwrap_or(0.0)
        .max(0.0)
        .mul_add(100.0, 0.0)
        .round() as i64
}
fn margin_basis_points(percent: &str) -> i64 {
    (percent
        .replace(',', ".")
        .parse::<f64>()
        .unwrap_or(0.0)
        .clamp(0.0, 10000.0)
        * 100.0)
        .round() as i64
}
fn sale_price(cost_cents: i64, margin_bps: i64) -> i64 {
    ((cost_cents.max(0) as i128 * (10_000 + margin_bps.clamp(0, 1_000_000)) as i128 + 5_000)
        / 10_000) as i64
}

#[derive(Clone)]
struct Reservation {
    id: i64,
    day: String,
    resource: String,
    customer: String,
    start: i64,
    duration: i64,
    games: String,
    total: i64,
    status: String,
    group_id: i64,
}
#[derive(Clone)]
struct Product {
    id: i64,
    name: String,
    price: i64,
    cost: i64,
    margin_bps: i64,
    stock: i64,
}
#[derive(Clone)]
struct Event {
    day: String,
    kind: String,
    title: String,
    customer: String,
}

#[derive(Clone)]
struct Promotion {
    id: i64,
    name: String,
    buy_qty: usize,
    pay_qty: usize,
    start_hour: i32,
    end_hour: i32,
    resource_kind: String,
}

#[derive(Clone)]
struct GameTimer {
    id: i64,
    resource: String,
    customer: String,
    started_at: i64,
    ends_at: i64,
    total_secs: i64,
}

#[derive(PartialEq, Clone, Copy)]
enum Page {
    Agenda,
    Reservations,
    Sales,
    Products,
    Prices,
    Promotions,
    Timers,
    Closing,
    Events,
}

struct VrBoxApp {
    db: Connection,
    page: Page,
    day: String,
    message: String,
    customer: String,
    selected_resources: BTreeSet<usize>,
    start_hour: i32,
    cart: BTreeMap<i64, i64>,
    product_name: String,
    product_cost: String,
    product_margin: String,
    product_stock: String,
    event_day: String,
    event_kind: String,
    event_title: String,
    event_customer: String,
    price_sim: String,
    price_gold: String,
    price_play: String,
    price_vr: String,
    counted_cash: String,
    counted_transfer: String,
    closing_note: String,
    pay_transfer: bool,
    session_game: String,
    promo_name: String,
    promo_buy: String,
    promo_pay: String,
    promo_start: i32,
    promo_end: i32,
    promo_kind: String,
    timer_resource: String,
    timer_customer: String,
    timer_minutes: String,
    editing_group_id: i64,
    editing_reservation_id: i64,
    reservation_payment_transfer: bool,
    reservation_partial_amount: String,
    pending_cancel: Option<(i64, i64)>,
}

impl VrBoxApp {
    fn new() -> Self {
        let db = open_db().expect("No se pudo iniciar la base local de VRBox");
        let day = Local::now().format("%Y-%m-%d").to_string();
        let price_sim = db
            .query_row(
                "SELECT value FROM settings WHERE key='price_sim_silver'",
                [],
                |r| r.get(0),
            )
            .unwrap_or_else(|_| "0".into());
        let price_gold = db
            .query_row(
                "SELECT value FROM settings WHERE key='price_sim_gold'",
                [],
                |r| r.get(0),
            )
            .unwrap_or_else(|_| "0".into());
        let price_play = db
            .query_row(
                "SELECT value FROM settings WHERE key='price_play'",
                [],
                |r| r.get(0),
            )
            .unwrap_or_else(|_| "0".into());
        let price_vr = db
            .query_row("SELECT value FROM settings WHERE key='price_vr'", [], |r| {
                r.get(0)
            })
            .unwrap_or_else(|_| "0".into());
        Self {
            db,
            page: Page::Agenda,
            day: day.clone(),
            message: "Datos guardados localmente en este equipo.".into(),
            customer: String::new(),
            selected_resources: BTreeSet::new(),
            start_hour: 16,
            cart: BTreeMap::new(),
            product_name: String::new(),
            product_cost: String::new(),
            product_margin: "50".into(),
            product_stock: "0".into(),
            event_day: day,
            event_kind: "Cumpleaños".into(),
            event_title: String::new(),
            event_customer: String::new(),
            price_sim,
            price_gold,
            price_play,
            price_vr,
            counted_cash: String::new(),
            counted_transfer: String::new(),
            closing_note: String::new(),
            pay_transfer: false,
            session_game: "Sin clasificar".into(),
            promo_name: String::new(),
            promo_buy: "2".into(),
            promo_pay: "1".into(),
            promo_start: 16,
            promo_end: 19,
            promo_kind: "all".into(),
            timer_resource: "PS5 1".into(),
            timer_customer: String::new(),
            timer_minutes: "30".into(),
            editing_group_id: 0,
            editing_reservation_id: 0,
            reservation_payment_transfer: false,
            reservation_partial_amount: String::new(),
            pending_cancel: None,
        }
    }
    fn price_for_kind(&self, kind: &str) -> i64 {
        let key = format!("price_{kind}");
        self.db
            .query_row("SELECT value FROM settings WHERE key=?1", [key], |r| {
                r.get::<_, String>(0)
            })
            .ok()
            .and_then(|v| v.parse().ok())
            .map(|v: i64| v.saturating_mul(100))
            .unwrap_or(0)
    }
    fn promotions(&self) -> Vec<Promotion> {
        let mut stmt = match self.db.prepare("SELECT id,name,buy_qty,pay_qty,start_hour,end_hour,resource_kind FROM promotions WHERE active=1 ORDER BY start_hour,name") {
            Ok(stmt) => stmt,
            Err(_) => return Vec::new(),
        };
        stmt.query_map([], |r| {
            Ok(Promotion {
                id: r.get(0)?,
                name: r.get(1)?,
                buy_qty: r.get::<_, i64>(2)? as usize,
                pay_qty: r.get::<_, i64>(3)? as usize,
                start_hour: r.get(4)?,
                end_hour: r.get(5)?,
                resource_kind: r.get(6)?,
            })
        })
        .map(|it| it.filter_map(Result::ok).collect())
        .unwrap_or_default()
    }
    fn game_timers(&self) -> Vec<GameTimer> {
        let mut stmt = match self.db.prepare("SELECT id,resource,customer,started_at,ends_at,total_secs FROM timers ORDER BY ends_at") { Ok(stmt) => stmt, Err(_) => return Vec::new() };
        stmt.query_map([], |r| {
            Ok(GameTimer {
                id: r.get(0)?,
                resource: r.get(1)?,
                customer: r.get(2)?,
                started_at: r.get(3)?,
                ends_at: r.get(4)?,
                total_secs: r.get(5)?,
            })
        })
        .map(|it| it.filter_map(Result::ok).collect())
        .unwrap_or_default()
    }
    fn add_timer(&mut self) {
        if !RESOURCES
            .iter()
            .any(|(name, kind)| *name == self.timer_resource && (*kind == "play" || *kind == "vr"))
        {
            self.message = "Elegí una estación PS5 o VR para iniciar el cronómetro.".into();
            return;
        }
        let minutes = self.timer_minutes.parse::<i64>().unwrap_or(0);
        if !(1..=1440).contains(&minutes) {
            self.message = "La duración debe ser de 1 a 1440 minutos.".into();
            return;
        }
        let now = Local::now().timestamp();
        let running = self
            .db
            .query_row(
                "SELECT COUNT(*) FROM timers WHERE resource=?1 AND ends_at>?2",
                params![self.timer_resource, now],
                |r| r.get::<_, i64>(0),
            )
            .unwrap_or(0);
        if running > 0 {
            self.message = format!("{} ya tiene un cronómetro en curso.", self.timer_resource);
            return;
        }
        let total = minutes * 60;
        match self.db.execute("INSERT INTO timers(resource,customer,started_at,ends_at,total_secs) VALUES(?1,?2,?3,?4,?5)", params![self.timer_resource, self.timer_customer.trim(), now, now + total, total]) {
            Ok(_) => { self.message = format!("Cronómetro iniciado para {}.", self.timer_resource); self.timer_customer.clear(); }
            Err(e) => self.message = format!("No se pudo iniciar el cronómetro: {e}"),
        }
    }
    fn adjust_timer(&mut self, id: i64, delta_minutes: i64) {
        let now = Local::now().timestamp();
        let result = self.db.execute("UPDATE timers SET ends_at=ends_at+?1,total_secs=MAX(60,total_secs+?1) WHERE id=?2 AND (?1>0 OR ends_at+?1>=?3)", params![delta_minutes * 60, id, now]);
        self.message = match result {
            Ok(1) => "Tiempo del cronómetro actualizado.".into(),
            Ok(_) => "No se pudo reducir más: el tiempo restante mínimo es 1 minuto.".into(),
            Err(e) => format!("No se pudo actualizar el cronómetro: {e}"),
        };
    }
    fn delete_timer(&mut self, id: i64) {
        self.message = match self.db.execute("DELETE FROM timers WHERE id=?1", [id]) {
            Ok(_) => "Cronómetro eliminado.".into(),
            Err(e) => format!("No se pudo eliminar: {e}"),
        };
    }
    /// Picks one best applicable promotion. For each complete bundle, the cheapest stations are free.
    fn reservation_quote(
        &self,
        selected: &BTreeSet<usize>,
        hour: i32,
    ) -> (Vec<(&'static str, i64, i64)>, String, i64) {
        let base = selected
            .iter()
            .map(|index| {
                (
                    RESOURCES[*index].0,
                    self.price_for_kind(RESOURCES[*index].1),
                    0_i64,
                )
            })
            .collect::<Vec<_>>();
        let mut best_name = String::new();
        let mut best_discount = 0_i64;
        let mut best_alloc = vec![0_i64; base.len()];
        for promo in self.promotions().into_iter().filter(|p| {
            p.start_hour <= hour && hour <= p.end_hour && p.buy_qty > p.pay_qty && p.buy_qty > 0
        }) {
            let mut eligible = base
                .iter()
                .enumerate()
                .filter(|(_, (resource, _, _))| {
                    selected.iter().any(|i| RESOURCES[*i].0 == *resource)
                        && (promo.resource_kind == "all"
                            || selected.iter().any(|i| {
                                RESOURCES[*i].0 == *resource
                                    && RESOURCES[*i].1 == promo.resource_kind
                            }))
                })
                .map(|(i, _)| i)
                .collect::<Vec<_>>();
            if promo.resource_kind != "all" {
                eligible.retain(|i| {
                    selected.iter().any(|index| {
                        RESOURCES[*index].0 == base[*i].0
                            && RESOURCES[*index].1 == promo.resource_kind
                    })
                });
            }
            eligible.sort_by_key(|i| (base[*i].1, base[*i].0));
            let free_count = eligible.len() / promo.buy_qty * (promo.buy_qty - promo.pay_qty);
            let mut alloc = vec![0_i64; base.len()];
            for index in eligible.into_iter().take(free_count) {
                alloc[index] = base[index].1;
            }
            let discount = alloc
                .iter()
                .fold(0_i64, |sum, amount| sum.saturating_add(*amount));
            if discount > best_discount {
                best_discount = discount;
                best_alloc = alloc;
                best_name = promo.name;
            }
        }
        let net = base
            .into_iter()
            .enumerate()
            .map(|(i, (name, price, _))| (name, price - best_alloc[i], best_alloc[i]))
            .collect();
        (net, best_name, best_discount)
    }
    fn reservations(&self, day: &str) -> Vec<Reservation> {
        let mut stmt = match self.db.prepare("SELECT id,service_date,resource,customer,start_hour,duration,games,total_cents,status,COALESCE(group_id,0) FROM reservations WHERE service_date=?1 ORDER BY start_hour,resource") { Ok(s) => s, Err(_) => return vec![] };
        stmt.query_map([day], |r| {
            Ok(Reservation {
                id: r.get(0)?,
                day: r.get(1)?,
                resource: r.get(2)?,
                customer: r.get(3)?,
                start: r.get(4)?,
                duration: r.get(5)?,
                games: r.get(6)?,
                total: r.get(7)?,
                status: r.get(8)?,
                group_id: r.get(9)?,
            })
        })
        .map(|it| it.filter_map(Result::ok).collect())
        .unwrap_or_default()
    }
    fn products(&self) -> Vec<Product> {
        let mut stmt = match self
            .db
            .prepare("SELECT id,name,price_cents,cost_cents,margin_bps,stock FROM products WHERE active=1 ORDER BY name")
        {
            Ok(s) => s,
            Err(_) => return vec![],
        };
        stmt.query_map([], |r| {
            Ok(Product {
                id: r.get(0)?,
                name: r.get(1)?,
                price: r.get(2)?,
                cost: r.get(3)?,
                margin_bps: r.get(4)?,
                stock: r.get(5)?,
            })
        })
        .map(|it| it.filter_map(Result::ok).collect())
        .unwrap_or_default()
    }
    fn events(&self) -> Vec<Event> {
        let mut stmt = match self.db.prepare("SELECT event_date,kind,title,customer FROM events WHERE event_date>=?1 ORDER BY event_date LIMIT 30") { Ok(s) => s, Err(_) => return vec![] };
        stmt.query_map([self.day.as_str()], |r| {
            Ok(Event {
                day: r.get(0)?,
                kind: r.get(1)?,
                title: r.get(2)?,
                customer: r.get(3)?,
            })
        })
        .map(|it| it.filter_map(Result::ok).collect())
        .unwrap_or_default()
    }
    fn create_reservation(&mut self) {
        let (stations, promo_name, discount) =
            self.reservation_quote(&self.selected_resources, self.start_hour);
        let day = self.day.clone();
        let customer = self.customer.clone();
        let station_count = stations.len();
        let editing_id = self.editing_group_id;
        let editing_single_id = self.editing_reservation_id;
        let result = if editing_id > 0 {
            reservations::update_group(
                &mut self.db,
                reservations::UpdateReservationGroup {
                    group_id: editing_id,
                    day: &day,
                    customer: &customer,
                    start_hour: self.start_hour,
                    promo_name: &promo_name,
                    stations: &stations,
                },
            )
            .map(|()| editing_id)
        } else if editing_single_id > 0 {
            if stations.len() != 1 {
                Err("Esta reserva histórica es individual: elegí exactamente una estación.".into())
            } else {
                let (station, price, line_discount) = stations[0];
                reservations::update_single(
                    &mut self.db,
                    reservations::UpdateSingleReservation {
                        reservation_id: editing_single_id,
                        day: &day,
                        customer: &customer,
                        start_hour: self.start_hour,
                        station,
                        total_cents: price,
                        discount_cents: line_discount,
                    },
                )
                .map(|()| editing_single_id)
            }
        } else {
            reservations::create_group(
                &mut self.db,
                reservations::NewReservationGroup {
                    day: &day,
                    customer: &customer,
                    start_hour: self.start_hour,
                    promo_name: &promo_name,
                    stations: &stations,
                },
            )
        };
        match result {
            Ok(group_id) => {
                self.customer.clear();
                self.selected_resources.clear();
                self.editing_group_id = 0;
                self.editing_reservation_id = 0;
                let total = stations
                    .iter()
                    .fold(0_i64, |sum, (_, amount, _)| sum.saturating_add(*amount));
                self.message = if editing_id > 0 || editing_single_id > 0 {
                    format!("Reserva #{group_id} modificada · {}.", money(total))
                } else if promo_name.is_empty() {
                    format!(
                        "Reserva grupal #{group_id}: {station_count} estación(es) · {}.",
                        money(total)
                    )
                } else {
                    format!(
                        "Reserva grupal #{group_id}: {promo_name} · descuento {} · total {}.",
                        money(discount),
                        money(total)
                    )
                };
            }
            Err(error) => self.message = error,
        }
    }
    fn begin_edit_group(&mut self, group_id: i64, stations: &[Reservation]) {
        if stations.is_empty() || stations.iter().any(|r| r.status != "reservada") {
            self.message = "Solo se pueden modificar grupos aún no iniciados.".into();
            return;
        }
        self.editing_group_id = group_id;
        self.editing_reservation_id = if group_id > 0 { 0 } else { stations[0].id };
        self.day = stations[0].day.clone();
        self.customer = stations[0].customer.clone();
        self.start_hour = stations[0].start as i32;
        self.selected_resources = stations
            .iter()
            .filter_map(|r| RESOURCES.iter().position(|(name, _)| *name == r.resource))
            .collect();
        self.page = Page::Reservations;
        let id = if group_id > 0 {
            group_id
        } else {
            stations[0].id
        };
        self.message =
            format!("Editando la reserva #{id}. Revisá el nuevo total antes de guardar.");
    }
    fn paid_for(&self, group_id: i64, reservation_id: i64) -> i64 {
        let query = if group_id > 0 {
            self.db.query_row("SELECT COALESCE(SUM(CASE WHEN kind='devolucion' THEN -amount_cents ELSE amount_cents END),0) FROM reservation_payments WHERE group_id=?1", [group_id], |r| r.get(0))
        } else {
            self.db.query_row("SELECT COALESCE(SUM(CASE WHEN kind='devolucion' THEN -amount_cents ELSE amount_cents END),0) FROM reservation_payments WHERE reservation_id=?1", [reservation_id], |r| r.get(0))
        };
        query.unwrap_or(0)
    }
    fn record_reservation_payment(
        &mut self,
        group_id: i64,
        reservation_id: i64,
        action: reservations::PaymentAction,
        requested: i64,
    ) {
        let method = if self.reservation_payment_transfer {
            "transferencia"
        } else {
            "efectivo"
        };
        let today = Local::now().format("%Y-%m-%d").to_string();
        match reservations::record_payment(
            &mut self.db,
            (group_id > 0).then_some(group_id),
            (reservation_id > 0).then_some(reservation_id),
            action,
            requested,
            method,
            &today,
        ) {
            Ok(amount) => {
                self.message = format!("Pago registrado · {} · {}.", method, money(amount))
            }
            Err(e) => self.message = e,
        }
    }
    fn cancel_reservation(&mut self, group_id: i64, reservation_id: i64) {
        let method = if self.reservation_payment_transfer {
            "transferencia"
        } else {
            "efectivo"
        };
        let today = Local::now().format("%Y-%m-%d").to_string();
        match reservations::cancel_group(
            &mut self.db,
            (group_id > 0).then_some(group_id),
            (reservation_id > 0).then_some(reservation_id),
            method,
            &today,
        ) {
            Ok(refund) if refund > 0 => {
                self.message = format!(
                    "Reserva cancelada · devolución registrada: {}.",
                    money(refund)
                )
            }
            Ok(_) => self.message = "Reserva cancelada.".into(),
            Err(e) => self.message = e,
        }
    }
    fn add_product(&mut self) {
        if self.product_name.trim().is_empty() {
            self.message = "Ingresá el nombre del producto.".into();
            return;
        }
        let cost = pesos_input(&self.product_cost);
        let margin = margin_basis_points(&self.product_margin);
        let price = sale_price(cost, margin);
        let stock = self.product_stock.parse::<i64>().unwrap_or(0).max(0);
        let result=self.db.execute("INSERT INTO products(name,price_cents,cost_cents,margin_bps,stock) VALUES(?1,?2,?3,?4,?5) ON CONFLICT(name) DO UPDATE SET price_cents=excluded.price_cents,cost_cents=excluded.cost_cents,margin_bps=excluded.margin_bps,stock=products.stock+excluded.stock,active=1",params![self.product_name.trim(),price,cost,margin,stock]);
        self.message = match result {
            Ok(_) => {
                self.product_name.clear();
                format!("Producto guardado · precio de venta {}.", money(price))
            }
            Err(e) => format!("Error: {e}"),
        };
    }
    fn add_to_cart(&mut self, product_id: i64) {
        if let Some(p) = self.products().into_iter().find(|p| p.id == product_id) {
            if p.stock <= *self.cart.get(&p.id).unwrap_or(&0) {
                self.message = format!("Stock disponible de {}: {}.", p.name, p.stock);
                return;
            }
            *self.cart.entry(p.id).or_default() += 1;
        }
    }
    fn checkout_cart(&mut self) {
        if self.cart.is_empty() {
            self.message = "El carrito está vacío.".into();
            return;
        }
        let payment = if self.pay_transfer {
            "transferencia"
        } else {
            "efectivo"
        };
        let products = self.products();
        let tx = match self
            .db
            .transaction_with_behavior(rusqlite::TransactionBehavior::Immediate)
        {
            Ok(tx) => tx,
            Err(e) => {
                self.message = format!("No se pudo abrir la venta: {e}");
                return;
            }
        };
        for (id, quantity) in &self.cart {
            let Some(p) = products.iter().find(|p| p.id == *id) else {
                self.message = "El catálogo cambió. Volvé a armar el carrito.".into();
                return;
            };
            if p.stock < *quantity {
                self.message = format!("Stock insuficiente para {}.", p.name);
                return;
            }
            let total = p.price * *quantity;
            let sold = tx.execute(
                "UPDATE products SET stock=stock-?1 WHERE id=?2 AND stock>=?1",
                params![quantity, p.id],
            );
            if !matches!(sold, Ok(1)) {
                self.message = format!(
                    "Stock insuficiente para {}. No se aplicó el carrito.",
                    p.name
                );
                return;
            }
            if let Err(e) = tx.execute("INSERT INTO sales(category,item,quantity,unit_cents,total_cents,payment,service_date) VALUES('bebida',?1,?2,?3,?4,?5,?6)", params![p.name, quantity, p.price, total, payment, self.day]) {
                self.message = format!("No se completó la venta: {e}");
                return;
            }
        }
        match tx.commit() {
            Ok(()) => {
                self.cart.clear();
                self.message = "Venta cobrada y stock actualizado.".into();
            }
            Err(e) => self.message = format!("No se completó la venta: {e}"),
        }
    }
    fn export_product_template(&mut self) {
        let Some(path) = rfd::FileDialog::new()
            .add_filter("Excel", &["xlsx"])
            .set_file_name("plantilla-productos-vrbox.xlsx")
            .save_file()
        else {
            return;
        };
        let mut book = Workbook::new();
        let head = Format::new().set_bold().set_background_color("#C9A75E");
        let sheet = book.add_worksheet();
        let _ = sheet.set_name("Productos");
        let headers = [
            "Producto",
            "Costo unitario (ARS)",
            "Recargo (%)",
            "Stock a sumar",
            "Precio de venta (ARS, guía)",
        ];
        for (col, title) in headers.iter().enumerate() {
            let _ = sheet.write_string(0, col as u16, *title);
            let _ = sheet.write_string(
                1,
                col as u16,
                if col == 0 { "Ejemplo: Gaseosa" } else { "" },
            );
        }
        let _ = sheet.set_row_format(0, &head);
        let _ = sheet.write_number(1, 1, 1000.0);
        let _ = sheet.write_number(1, 2, 50.0);
        let _ = sheet.write_number(1, 3, 12.0);
        let _ = sheet.write_formula(1, 4, Formula::new("=B2*(1+C2/100)"));
        for (col, width) in [28.0, 24.0, 16.0, 18.0, 30.0].iter().enumerate() {
            let _ = sheet.set_column_width(col as u16, *width);
        }
        let guide = book.add_worksheet();
        let _ = guide.set_name("Instrucciones");
        let instructions = [
            "Cómo importar productos en VRBox",
            "Completá la hoja Productos desde la fila 2.",
            "Producto: nombre único. Costo unitario y Recargo (%) determinan el precio de venta.",
            "Stock a sumar: unidades compradas; importar el mismo producto repone stock.",
            "La columna Precio de venta es solo una guía calculada. VRBox la recalcula al importar.",
            "No cambies los títulos de las columnas. Guardá el archivo como .xlsx.",
        ];
        for (row, line) in instructions.iter().enumerate() {
            let _ = guide.write_string(row as u32, 0, *line);
        }
        let _ = guide.set_column_width(0, 100.0);
        self.message = match book.save(&path) {
            Ok(()) => format!("Plantilla guardada en {}", path.display()),
            Err(e) => format!("No se pudo guardar la plantilla: {e}"),
        };
    }
    fn import_products(&mut self) {
        let Some(path) = rfd::FileDialog::new()
            .add_filter("Excel", &["xlsx"])
            .pick_file()
        else {
            return;
        };
        let parsed = (|| -> Result<Vec<(String, i64, i64, i64)>, String> {
            let mut workbook = open_workbook_auto(&path).map_err(|e| e.to_string())?;
            let name = workbook
                .sheet_names()
                .first()
                .cloned()
                .ok_or("El libro no contiene hojas")?;
            let range = workbook.worksheet_range(&name).map_err(|e| e.to_string())?;
            let mut rows = range.rows();
            let header = rows
                .next()
                .ok_or("Falta la fila de encabezados")?
                .iter()
                .map(|v| v.to_string().trim().to_lowercase())
                .collect::<Vec<_>>();
            let find_col = |needle: &str| {
                header
                    .iter()
                    .position(|h| h == needle)
                    .ok_or_else(|| format!("Falta la columna '{needle}'"))
            };
            let product_col = find_col("producto")?;
            let cost_col = find_col("costo unitario (ars)")?;
            let margin_col = find_col("recargo (%)")?;
            let stock_col = find_col("stock a sumar")?;
            let mut result = Vec::new();
            for (index, row) in rows.enumerate() {
                let value = |column: usize| {
                    row.get(column)
                        .map(ToString::to_string)
                        .unwrap_or_default()
                        .trim()
                        .to_string()
                };
                let product = value(product_col);
                if product.is_empty() || product.to_lowercase().starts_with("ejemplo:") {
                    continue;
                }
                let parse_number = |s: String| {
                    s.replace('$', "")
                        .trim()
                        .parse::<f64>()
                        .map_err(|_| format!("Fila {}: revisá los números", index + 2))
                };
                let cost = parse_number(value(cost_col))?;
                let margin = parse_number(value(margin_col))?;
                let stock = parse_number(value(stock_col))?;
                if cost < 0.0 || margin < 0.0 || stock < 0.0 {
                    return Err(format!(
                        "Fila {}: no se permiten valores negativos",
                        index + 2
                    ));
                }
                let cost_cents = (cost * 100.0).round() as i64;
                let margin_bps = (margin * 100.0).round() as i64;
                result.push((product, cost_cents, margin_bps, stock as i64));
            }
            if result.is_empty() {
                return Err("No hay productos cargados para importar".into());
            }
            Ok(result)
        })();
        let rows = match parsed {
            Ok(rows) => rows,
            Err(e) => {
                self.message = format!("No se importó ningún producto: {e}");
                return;
            }
        };
        let tx = match self
            .db
            .transaction_with_behavior(rusqlite::TransactionBehavior::Immediate)
        {
            Ok(tx) => tx,
            Err(e) => {
                self.message = format!("No se pudo iniciar la importación: {e}");
                return;
            }
        };
        for (name, cost, margin, stock) in &rows {
            let price = sale_price(*cost, *margin);
            if let Err(e) = tx.execute("INSERT INTO products(name,price_cents,cost_cents,margin_bps,stock) VALUES(?1,?2,?3,?4,?5) ON CONFLICT(name) DO UPDATE SET price_cents=excluded.price_cents,cost_cents=excluded.cost_cents,margin_bps=excluded.margin_bps,stock=products.stock+excluded.stock,active=1", params![name,price,cost,margin,stock]) {
                self.message = format!("Importación cancelada sin aplicar cambios: {e}"); return;
            }
        }
        match tx.commit() {
            Ok(()) => {
                self.message = format!(
                    "{} productos importados. El stock se sumó al existente.",
                    rows.len()
                )
            }
            Err(e) => self.message = format!("No se pudo completar la importación: {e}"),
        }
    }
    fn save_prices(&mut self) {
        for (key, value) in [
            ("price_sim_silver", &self.price_sim),
            ("price_sim_gold", &self.price_gold),
            ("price_play", &self.price_play),
            ("price_vr", &self.price_vr),
        ] {
            let _=self.db.execute("INSERT INTO settings(key,value) VALUES(?1,?2) ON CONFLICT(key) DO UPDATE SET value=excluded.value",params![key,value.parse::<i64>().unwrap_or(0).clamp(0,1_000_000_000).to_string()]);
        }
        self.message = "Tarifas guardadas. Se aplican a nuevas reservas.".into();
    }
    fn add_promotion(&mut self) {
        let buy = self.promo_buy.parse::<i64>().unwrap_or(0);
        let pay = self.promo_pay.parse::<i64>().unwrap_or(0);
        if self.promo_name.trim().is_empty() || buy < 2 || pay < 1 || pay >= buy {
            self.message =
                "Completá el nombre y una mecánica válida (por ejemplo, 2 por 1).".into();
            return;
        }
        if !(16..=23).contains(&self.promo_start)
            || !(self.promo_start..=23).contains(&self.promo_end)
        {
            self.message = "El horario de la promoción debe estar entre 16:00 y 23:00.".into();
            return;
        }
        match self.db.execute("INSERT INTO promotions(name,buy_qty,pay_qty,start_hour,end_hour,resource_kind) VALUES(?1,?2,?3,?4,?5,?6)", params![self.promo_name.trim(), buy, pay, self.promo_start, self.promo_end, self.promo_kind]) {
            Ok(_) => { self.message = format!("Promoción '{}' guardada.", self.promo_name.trim()); self.promo_name.clear(); }
            Err(e) => self.message = format!("No se pudo guardar la promoción: {e}"),
        }
    }
    fn add_event(&mut self) {
        if self.event_title.trim().is_empty() {
            self.message = "Ingresá el título del evento.".into();
            return;
        }
        let result = self.db.execute(
            "INSERT INTO events(event_date,kind,title,customer) VALUES(?1,?2,?3,?4)",
            params![
                self.event_day,
                self.event_kind,
                self.event_title.trim(),
                self.event_customer.trim()
            ],
        );
        self.message = if result.is_ok() {
            self.event_title.clear();
            "Evento guardado.".into()
        } else {
            "No se pudo guardar el evento.".into()
        };
    }
    fn totals(&self) -> (i64, i64, i64, i64) {
        let row=self.db.query_row("SELECT COALESCE(SUM(CASE WHEN payment='efectivo' THEN total_cents ELSE 0 END),0), COALESCE(SUM(CASE WHEN payment='transferencia' THEN total_cents ELSE 0 END),0), COALESCE(SUM(CASE WHEN category='bebida' THEN total_cents ELSE 0 END),0), COUNT(*) FROM sales WHERE service_date=?1",[self.day.as_str()],|r|Ok((r.get(0)?,r.get(1)?,r.get(2)?,r.get(3)?))).unwrap_or((0,0,0,0));
        let reserved_groups: i64 = self.db.query_row(
            "SELECT COALESCE(SUM(MAX(0,g.total_cents-COALESCE((SELECT SUM(CASE WHEN p.kind='devolucion' THEN -p.amount_cents ELSE p.amount_cents END) FROM reservation_payments p WHERE p.group_id=g.id),0))),0) FROM reservation_groups g WHERE g.service_date=?1 AND EXISTS(SELECT 1 FROM reservations r WHERE r.group_id=g.id AND r.status='reservada')",
            [self.day.as_str()], |r| r.get(0)
        ).unwrap_or(0);
        let reserved_single: i64 = self.db.query_row(
            "SELECT COALESCE(SUM(MAX(0,r.total_cents-COALESCE((SELECT SUM(CASE WHEN p.kind='devolucion' THEN -p.amount_cents ELSE p.amount_cents END) FROM reservation_payments p WHERE p.reservation_id=r.id),0))),0) FROM reservations r WHERE r.group_id IS NULL AND r.service_date=?1 AND r.status='reservada'",
            [self.day.as_str()], |r| r.get(0)
        ).unwrap_or(0);
        let reserved = reserved_groups.saturating_add(reserved_single);
        (row.0, row.1, row.2, reserved)
    }
    fn register_session(&mut self, r: &Reservation) {
        let game = self.session_game.clone();
        let tx = match self.db.transaction() {
            Ok(tx) => tx,
            Err(e) => {
                self.message = format!("No se pudo iniciar el cobro: {e}");
                return;
            }
        };
        let changed = tx.execute("UPDATE reservations SET status='completada',games=?1 WHERE id=?2 AND status='reservada'",params![game,r.id]);
        if !matches!(changed, Ok(1)) {
            self.message = "Esta estación ya se cobró o cambió de estado. Actualizá la agenda antes de continuar.".into();
            return;
        }
        match tx.commit() {
            Ok(()) => {
                self.message =
                    format!("Sesión completada · {game}. Consultá el estado del pago en Reservas.");
            }
            Err(e) => self.message = format!("No se pudo completar la sesión: {e}"),
        }
    }
    fn close_and_export(&mut self) {
        let (cash, transfer, _, reserved) = self.totals();
        let counted_cash = pesos_input(&self.counted_cash);
        let counted_transfer = pesos_input(&self.counted_transfer);
        let result=self.db.execute("INSERT INTO closings(service_date,expected_cash,expected_transfer,counted_cash,counted_transfer,notes) VALUES(?1,?2,?3,?4,?5,?6) ON CONFLICT(service_date) DO UPDATE SET expected_cash=excluded.expected_cash,expected_transfer=excluded.expected_transfer,counted_cash=excluded.counted_cash,counted_transfer=excluded.counted_transfer,closed_at=CURRENT_TIMESTAMP,notes=excluded.notes",params![self.day,cash,transfer,counted_cash,counted_transfer,self.closing_note]);
        if let Err(e) = result {
            self.message = format!("No se pudo guardar el cierre: {e}");
            return;
        }
        match self.export_xlsx(cash, transfer, reserved, counted_cash, counted_transfer) {
            Ok(path) => {
                self.message = format!("Cierre guardado y Excel exportado: {}", path.display())
            }
            Err(e) => self.message = format!("Cierre guardado, pero falló el Excel: {e}"),
        }
    }
    fn export_xlsx(
        &self,
        cash: i64,
        transfer: i64,
        reserved: i64,
        counted_cash: i64,
        counted_transfer: i64,
    ) -> Result<PathBuf, String> {
        let mut workbook = Workbook::new();
        let header = Format::new().set_bold().set_background_color("#D7B56D");
        {
            let sheet = workbook.add_worksheet();
            sheet.set_name("Resumen").map_err(|e| e.to_string())?;
            for (c, h) in ["VRBox · cierre diario", "Valor"].iter().enumerate() {
                sheet
                    .write_string(0, c as u16, *h)
                    .map_err(|e| e.to_string())?;
            }
            for (row, (k, v)) in [
                ("Fecha", self.day.clone()),
                ("Efectivo esperado", money(cash)),
                ("Transferencias", money(transfer)),
                ("Efectivo contado", money(counted_cash)),
                ("Transferencia informada", money(counted_transfer)),
                ("Sesiones reservadas pendientes", money(reserved)),
            ]
            .iter()
            .enumerate()
            {
                sheet
                    .write_string(row as u32 + 1, 0, *k)
                    .map_err(|e| e.to_string())?;
                sheet
                    .write_string(row as u32 + 1, 1, v)
                    .map_err(|e| e.to_string())?;
            }
            sheet.set_column_width(0, 32).map_err(|e| e.to_string())?;
            sheet.set_column_width(1, 28).map_err(|e| e.to_string())?;
            sheet
                .set_row_format(0, &header)
                .map_err(|e| e.to_string())?;
        }
        {
            let sheet = workbook.add_worksheet();
            sheet.set_name("Ventas").map_err(|e| e.to_string())?;
            let headers = [
                "Hora",
                "Tipo",
                "Concepto",
                "Cantidad",
                "Unitario centavos",
                "Total centavos",
                "Medio",
            ];
            for (c, h) in headers.iter().enumerate() {
                sheet
                    .write_string(0, c as u16, *h)
                    .map_err(|e| e.to_string())?;
            }
            let mut stmt=self.db.prepare("SELECT sold_at,category,item,quantity,unit_cents,total_cents,payment FROM sales WHERE service_date=?1 ORDER BY sold_at").map_err(|e|e.to_string())?;
            let rows = stmt
                .query_map([self.day.as_str()], |r| {
                    Ok((
                        r.get::<_, String>(0)?,
                        r.get::<_, String>(1)?,
                        r.get::<_, String>(2)?,
                        r.get::<_, i64>(3)?,
                        r.get::<_, i64>(4)?,
                        r.get::<_, i64>(5)?,
                        r.get::<_, String>(6)?,
                    ))
                })
                .map_err(|e| e.to_string())?;
            for (i, row) in rows.enumerate() {
                let (a, b, c, d, e, f, g) = row.map_err(|e| e.to_string())?;
                let r = i as u32 + 1;
                for (col, val) in [(0, a), (1, b), (2, c), (6, g)] {
                    sheet.write_string(r, col, val).map_err(|e| e.to_string())?;
                }
                for (col, val) in [(3, d), (4, e), (5, f)] {
                    sheet
                        .write_number(r, col, val as f64)
                        .map_err(|e| e.to_string())?;
                }
            }
            for col in 0..7 {
                sheet.set_column_width(col, 20).map_err(|e| e.to_string())?;
            }
        }
        {
            let sheet = workbook.add_worksheet();
            sheet
                .set_name("Pagos reservas")
                .map_err(|e| e.to_string())?;
            for (col, title) in [
                "Fecha y hora",
                "Cliente",
                "Tipo",
                "Monto centavos",
                "Medio",
                "Reserva",
            ]
            .iter()
            .enumerate()
            {
                sheet
                    .write_string(0, col as u16, *title)
                    .map_err(|e| e.to_string())?;
            }
            let mut stmt = self.db.prepare(
                "SELECT p.paid_at,COALESCE(g.customer,r.customer,''),p.kind,p.amount_cents,p.payment,COALESCE('Grupo #'||p.group_id,'Reserva #'||p.reservation_id) FROM reservation_payments p LEFT JOIN reservation_groups g ON g.id=p.group_id LEFT JOIN reservations r ON r.id=p.reservation_id WHERE p.service_date=?1 ORDER BY p.paid_at,p.id"
            ).map_err(|e| e.to_string())?;
            let rows = stmt
                .query_map([self.day.as_str()], |r| {
                    Ok((
                        r.get::<_, String>(0)?,
                        r.get::<_, String>(1)?,
                        r.get::<_, String>(2)?,
                        r.get::<_, i64>(3)?,
                        r.get::<_, String>(4)?,
                        r.get::<_, String>(5)?,
                    ))
                })
                .map_err(|e| e.to_string())?;
            for (i, row) in rows.enumerate() {
                let (at, customer, kind, amount, method, target) =
                    row.map_err(|e| e.to_string())?;
                let row = i as u32 + 1;
                for (col, value) in [(0, at), (1, customer), (2, kind), (4, method), (5, target)] {
                    sheet
                        .write_string(row, col, value)
                        .map_err(|e| e.to_string())?;
                }
                sheet
                    .write_number(row, 3, amount as f64)
                    .map_err(|e| e.to_string())?;
            }
            for col in 0..6 {
                sheet.set_column_width(col, 24).map_err(|e| e.to_string())?;
            }
        }
        {
            let sheet = workbook.add_worksheet();
            sheet.set_name("Agenda").map_err(|e| e.to_string())?;
            for (c, h) in [
                "Fecha",
                "Recurso",
                "Cliente",
                "Inicio",
                "Duración h",
                "Juegos",
                "Total centavos",
                "Estado",
                "Grupo reserva",
            ]
            .iter()
            .enumerate()
            {
                sheet
                    .write_string(0, c as u16, *h)
                    .map_err(|e| e.to_string())?;
            }
            for (i, r) in self.reservations(&self.day).iter().enumerate() {
                let row = i as u32 + 1;
                for (col, val) in [
                    (0, r.day.clone()),
                    (1, r.resource.clone()),
                    (2, r.customer.clone()),
                    (5, r.games.clone()),
                    (7, r.status.clone()),
                    (8, r.group_id.to_string()),
                ] {
                    sheet
                        .write_string(row, col, val)
                        .map_err(|e| e.to_string())?;
                }
                for (col, val) in [(3, r.start), (4, r.duration), (6, r.total)] {
                    sheet
                        .write_number(row, col, val as f64)
                        .map_err(|e| e.to_string())?;
                }
            }
            for col in 0..9 {
                sheet.set_column_width(col, 21).map_err(|e| e.to_string())?;
            }
        }
        {
            let sheet = workbook.add_worksheet();
            sheet.set_name("Por juego").map_err(|e| e.to_string())?;
            for (col, title) in [
                "Juego",
                "Estaciones completadas",
                "Total por juego centavos",
            ]
            .iter()
            .enumerate()
            {
                sheet
                    .write_string(0, col as u16, *title)
                    .map_err(|e| e.to_string())?;
            }
            let mut totals: BTreeMap<String, (i64, i64)> = BTreeMap::new();
            let mut stmt = self.db.prepare("SELECT games,total_cents FROM reservations WHERE service_date=?1 AND status='completada' ORDER BY id").map_err(|e| e.to_string())?;
            let rows = stmt
                .query_map([self.day.as_str()], |r| {
                    Ok((r.get::<_, String>(0)?, r.get::<_, i64>(1)?))
                })
                .map_err(|e| e.to_string())?;
            for row in rows {
                let (games_text, total) = row.map_err(|e| e.to_string())?;
                let games = games_text
                    .split(", ")
                    .filter(|g| !g.is_empty())
                    .collect::<Vec<_>>();
                if games.is_empty() {
                    continue;
                }
                let each = total / games.len() as i64;
                let remainder = total % games.len() as i64;
                for (index, game) in games.iter().enumerate() {
                    let attributed = each + i64::from((index as i64) < remainder);
                    let entry = totals.entry((*game).to_string()).or_default();
                    entry.0 += 1;
                    entry.1 += attributed;
                }
            }
            sheet
                .write_string(
                    1,
                    0,
                    "Uso completado por juego. Los cobros por medio están en la hoja Ventas.",
                )
                .map_err(|e| e.to_string())?;
            for (index, (game, (sessions, total))) in totals.iter().enumerate() {
                let row = index as u32 + 2;
                sheet
                    .write_string(row, 0, game)
                    .map_err(|e| e.to_string())?;
                sheet
                    .write_number(row, 1, *sessions as f64)
                    .map_err(|e| e.to_string())?;
                sheet
                    .write_number(row, 2, *total as f64)
                    .map_err(|e| e.to_string())?;
            }
            for col in 0..4 {
                sheet.set_column_width(col, 30).map_err(|e| e.to_string())?;
            }
        }
        {
            let sheet = workbook.add_worksheet();
            sheet
                .set_name("Horas vendidas")
                .map_err(|e| e.to_string())?;
            sheet
                .write_string(0, 0, "Franja")
                .map_err(|e| e.to_string())?;
            sheet
                .write_string(0, 1, "Horas de estación cobradas")
                .map_err(|e| e.to_string())?;
            let mut occupancy = [0_i64; 8];
            for r in self
                .reservations(&self.day)
                .iter()
                .filter(|r| r.status == "completada")
            {
                for hour in r.start..(r.start + r.duration) {
                    if (16..24).contains(&hour) {
                        occupancy[(hour - 16) as usize] += 1;
                    }
                }
            }
            for (i, count) in occupancy.iter().enumerate() {
                sheet
                    .write_string(
                        i as u32 + 1,
                        0,
                        format!("{:02}:00–{:02}:00", i + 16, i + 17),
                    )
                    .map_err(|e| e.to_string())?;
                sheet
                    .write_number(i as u32 + 1, 1, *count as f64)
                    .map_err(|e| e.to_string())?;
            }
            sheet.set_column_width(0, 24).map_err(|e| e.to_string())?;
            sheet.set_column_width(1, 32).map_err(|e| e.to_string())?;
        }
        {
            let sheet = workbook.add_worksheet();
            sheet.set_name("Futuros").map_err(|e| e.to_string())?;
            for (col, title) in [
                "Tipo",
                "Fecha",
                "Recurso / categoría",
                "Cliente / evento",
                "Horario",
                "Importe centavos",
            ]
            .iter()
            .enumerate()
            {
                sheet
                    .write_string(0, col as u16, *title)
                    .map_err(|e| e.to_string())?;
            }
            let mut row = 1_u32;
            let mut reservations = self.db.prepare("SELECT service_date,resource,customer,start_hour,duration,total_cents FROM reservations WHERE service_date>?1 AND status='reservada' ORDER BY service_date,start_hour").map_err(|e| e.to_string())?;
            let future_reservations = reservations
                .query_map([self.day.as_str()], |r| {
                    Ok((
                        r.get::<_, String>(0)?,
                        r.get::<_, String>(1)?,
                        r.get::<_, String>(2)?,
                        r.get::<_, i64>(3)?,
                        r.get::<_, i64>(4)?,
                        r.get::<_, i64>(5)?,
                    ))
                })
                .map_err(|e| e.to_string())?;
            for item in future_reservations {
                let (day, resource, customer, start, duration, amount) =
                    item.map_err(|e| e.to_string())?;
                for (col, value) in [
                    (0, "Reserva".to_string()),
                    (1, day),
                    (2, resource),
                    (3, customer),
                    (4, format!("{:02}:00 · {} h", start, duration)),
                ] {
                    sheet
                        .write_string(row, col, value)
                        .map_err(|e| e.to_string())?;
                }
                sheet
                    .write_number(row, 5, amount as f64)
                    .map_err(|e| e.to_string())?;
                row += 1;
            }
            let mut events = self.db.prepare("SELECT event_date,kind,title,customer FROM events WHERE event_date>=?1 ORDER BY event_date").map_err(|e|e.to_string())?;
            let future_events = events
                .query_map([self.day.as_str()], |r| {
                    Ok((
                        r.get::<_, String>(0)?,
                        r.get::<_, String>(1)?,
                        r.get::<_, String>(2)?,
                        r.get::<_, String>(3)?,
                    ))
                })
                .map_err(|e| e.to_string())?;
            for item in future_events {
                let (day, kind, title, customer) = item.map_err(|e| e.to_string())?;
                for (col, value) in [(0, kind), (1, day), (2, title), (3, customer)] {
                    sheet
                        .write_string(row, col, value)
                        .map_err(|e| e.to_string())?;
                }
                row += 1;
            }
            for col in 0..6 {
                sheet.set_column_width(col, 28).map_err(|e| e.to_string())?;
            }
        }
        let path = data_dir().join(format!("cierre-{}.xlsx", self.day));
        workbook.save(&path).map_err(|e| e.to_string())?;
        Ok(path)
    }
}

impl eframe::App for VrBoxApp {
    fn ui(&mut self, root: &mut egui::Ui, _frame: &mut eframe::Frame) {
        let ctx = root.ctx().clone();
        ctx.request_repaint_after(Duration::from_secs(1));
        let now = Local::now().timestamp();
        let all_timers = self.game_timers();
        let expired_timers = all_timers
            .iter()
            .filter(|timer| timer.ends_at <= now)
            .cloned()
            .collect::<Vec<_>>();
        let mut visuals = egui::Visuals::dark();
        visuals.panel_fill = Color32::from_rgb(10, 15, 27);
        visuals.window_fill = Color32::from_rgb(17, 25, 41);
        visuals.extreme_bg_color = Color32::from_rgb(13, 21, 36);
        visuals.faint_bg_color = Color32::from_rgb(24, 35, 54);
        visuals.selection.bg_fill = Color32::from_rgb(26, 146, 167);
        visuals.widgets.inactive.bg_fill = Color32::from_rgb(25, 37, 57);
        visuals.widgets.inactive.weak_bg_fill = Color32::from_rgb(25, 37, 57);
        visuals.widgets.inactive.bg_stroke.color = Color32::from_rgb(48, 66, 88);
        visuals.widgets.hovered.bg_fill = Color32::from_rgb(33, 63, 82);
        visuals.widgets.hovered.weak_bg_fill = Color32::from_rgb(33, 63, 82);
        visuals.widgets.active.bg_fill = Color32::from_rgb(28, 116, 135);
        for widget in [
            &mut visuals.widgets.inactive,
            &mut visuals.widgets.hovered,
            &mut visuals.widgets.active,
        ] {
            widget.corner_radius = egui::CornerRadius::same(10);
        }
        ctx.set_visuals(visuals);
        let mut style = (*ctx.global_style()).clone();
        style.animation_time = 0.22;
        style.spacing.item_spacing = Vec2::new(10.0, 10.0);
        style.spacing.button_padding = Vec2::new(13.0, 9.0);
        ctx.set_global_style(style);
        egui::Panel::top("top").show_inside(root, |ui| {
            egui::Frame::new()
                .fill(Color32::from_rgb(13, 20, 34))
                .inner_margin(egui::Margin::symmetric(18, 12))
                .show(ui, |ui| {
                    ui.horizontal(|ui| {
                        ui.heading(
                            RichText::new("VRBOX")
                                .color(Color32::from_rgb(231, 195, 107))
                                .strong(),
                        );
                        ui.label(
                            RichText::new("NEXUS  /  OPERACIONES")
                                .small()
                                .color(Color32::from_rgb(107, 190, 203)),
                        );
                        ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                            if let Some(timer) = expired_timers.first() {
                                if ui
                                    .button(
                                        RichText::new(format!(
                                            "TIEMPO FINALIZADO · {}",
                                            timer.resource
                                        ))
                                        .color(Color32::from_rgb(255, 151, 142)),
                                    )
                                    .clicked()
                                {
                                    self.page = Page::Timers;
                                }
                            } else if !all_timers.is_empty() {
                                ui.label(
                                    RichText::new("CRONÓMETROS EN CURSO")
                                        .small()
                                        .color(Color32::from_rgb(91, 203, 196)),
                                );
                            }
                            ui.label(
                                RichText::new("FECHA OPERATIVA")
                                    .small()
                                    .color(Color32::GRAY),
                            );
                            ui.add_sized([130.0, 30.0], egui::TextEdit::singleline(&mut self.day));
                            ui.label(
                                RichText::new("SISTEMA LOCAL")
                                    .color(Color32::from_rgb(107, 206, 166))
                                    .small(),
                            );
                        });
                    });
                });
        });
        egui::Panel::left("navigation")
            .min_size(205.0)
            .show_inside(root, |ui| {
                ui.add_space(18.0);
                ui.label(
                    RichText::new("ESPACIO DE TRABAJO")
                        .small()
                        .color(Color32::from_rgb(102, 127, 152)),
                );
                ui.add_space(10.0);
                for (page, icon, label) in [
                    (Page::Agenda, "AG", "Agenda"),
                    (Page::Reservations, "RS", "Reservas"),
                    (Page::Sales, "VT", "Punto de venta"),
                    (Page::Products, "ST", "Bebidas y stock"),
                    (Page::Prices, "$", "Tarifas"),
                    (Page::Promotions, "OF", "Promociones"),
                    (Page::Timers, "CR", "Cronómetros"),
                    (Page::Closing, "CJ", "Cierre de caja"),
                    (Page::Events, "EV", "Eventos y alertas"),
                ] {
                    let selected = self.page == page;
                    let transition =
                        ctx.animate_bool_with_time(egui::Id::new(("nav", label)), selected, 0.20);
                    let fill = Color32::from_rgb(10, 15, 27)
                        .lerp_to_gamma(Color32::from_rgb(27, 48, 70), transition);
                    let text_color = if selected {
                        Color32::from_rgb(230, 192, 105)
                    } else {
                        Color32::from_rgb(188, 204, 222)
                    };
                    let response = egui::Frame::new()
                        .fill(fill)
                        .corner_radius(egui::CornerRadius::same(10))
                        .inner_margin(egui::Margin::symmetric(10, 11))
                        .show(ui, |ui| {
                            ui.set_min_width(164.0);
                            ui.horizontal(|ui| {
                                ui.add_sized(
                                    [28.0, 24.0],
                                    egui::Label::new(
                                        RichText::new(icon).color(text_color).monospace().strong(),
                                    )
                                    .selectable(false),
                                );
                                ui.label(RichText::new(label).color(text_color).strong());
                            });
                        })
                        .response
                        .interact(egui::Sense::click());
                    if response.hovered() {
                        ui.ctx().set_cursor_icon(egui::CursorIcon::PointingHand);
                    }
                    if response.clicked() {
                        self.page = page;
                    }
                }
                ui.separator();
                ui.add_space(8.0);
                ui.label(
                    RichText::new("VRBOX  •  LOCAL")
                        .small()
                        .color(Color32::from_rgb(103, 127, 151)),
                );
                ui.small("Datos privados en este equipo");
            });
        egui::CentralPanel::default().show_inside(root, |ui| {
            ui.add_space(4.0);
            egui::ScrollArea::vertical()
                .auto_shrink([false, false])
                .show(ui, |ui| {
                    egui::Frame::new()
                        .fill(Color32::from_rgb(14, 22, 37))
                        .stroke(egui::Stroke::new(1.0_f32, Color32::from_rgb(34, 49, 70)))
                        .corner_radius(egui::CornerRadius::same(16))
                        .inner_margin(egui::Margin::same(20))
                        .show(ui, |ui| {
                            ui.set_min_width((ui.available_width() - 40.0).max(600.0));
                            match self.page {
                                Page::Agenda => self.ui_agenda(ui),
                                Page::Reservations => self.ui_reservations(ui),
                                Page::Sales => self.ui_sales(ui),
                                Page::Products => self.ui_products(ui),
                                Page::Prices => self.ui_prices(ui),
                                Page::Promotions => self.ui_promotions(ui),
                                Page::Timers => self.ui_timers(ui),
                                Page::Closing => self.ui_closing(ui),
                                Page::Events => self.ui_events(ui),
                            }
                            ui.separator();
                            ui.label(
                                RichText::new(&self.message)
                                    .color(Color32::from_rgb(143, 175, 194)),
                            );
                        });
                });
        });
        if let Some((group_id, reservation_id)) = self.pending_cancel {
            let paid = self.paid_for(group_id, reservation_id);
            let method = if self.reservation_payment_transfer {
                "transferencia"
            } else {
                "efectivo"
            };
            let mut confirm = false;
            let mut dismiss = false;
            egui::Window::new("Confirmar cancelación")
                .collapsible(false)
                .resizable(false)
                .show(&ctx, |ui| {
                    ui.label("La cancelación quedará registrada en el historial.");
                    if paid > 0 {
                        ui.label(format!(
                            "Se devolverán {} por {} y se registrará la salida en caja.",
                            money(paid),
                            method
                        ));
                    } else {
                        ui.label("No hay pagos asociados a esta reserva.");
                    }
                    ui.horizontal(|ui| {
                        if ui.button("Volver").clicked() {
                            dismiss = true;
                        }
                        if ui.button("Confirmar cancelación").clicked() {
                            confirm = true;
                        }
                    });
                });
            if confirm {
                self.pending_cancel = None;
                self.cancel_reservation(group_id, reservation_id);
            } else if dismiss {
                self.pending_cancel = None;
            }
        }
        ctx.request_repaint_after(Duration::from_secs(1));
    }
}

impl VrBoxApp {
    fn ui_agenda(&mut self, ui: &mut egui::Ui) {
        ui.label(
            RichText::new("AGENDA  /  SALA PRINCIPAL")
                .small()
                .color(Color32::from_rgb(91, 186, 200)),
        );
        ui.heading(
            RichText::new("Tu jornada, de un vistazo")
                .size(27.0)
                .strong(),
        );
        ui.label(
            RichText::new(format!("{}  ·  16:00 — 00:00", self.day))
                .color(Color32::from_rgb(157, 176, 197)),
        );
        let rs = self.reservations(&self.day);
        let (cash, transfer, _, pending) = self.totals();
        ui.horizontal_wrapped(|ui| {
            stat(ui, "Efectivo registrado", money(cash));
            stat(ui, "Transferencias", money(transfer));
            stat(ui, "Reservas pendientes", money(pending));
        });
        ui.add_space(10.0);
        egui::ScrollArea::both().show(ui, |ui| {
            egui::Grid::new("agenda_grid")
                .striped(false)
                .min_col_width(106.0)
                .spacing(Vec2::new(7.0, 7.0))
                .show(ui, |ui| {
                    ui.add_sized(
                        [84.0, 42.0],
                        egui::Label::new(RichText::new("HORA").small().color(Color32::GRAY)),
                    );
                    for (idx, (name, _)) in RESOURCES.iter().enumerate() {
                        let color = if idx >= 6 {
                            Color32::from_rgb(231, 190, 92)
                        } else {
                            Color32::from_rgb(151, 180, 208)
                        };
                        ui.add_sized(
                            [106.0, 42.0],
                            egui::Label::new(
                                RichText::new(if idx >= 6 {
                                    name.to_string()
                                } else {
                                    (*name).to_string()
                                })
                                .color(color)
                                .strong(),
                            )
                            .wrap(),
                        );
                    }
                    ui.end_row();
                    for hour in 16..24 {
                        ui.add_sized(
                            [84.0, 64.0],
                            egui::Label::new(
                                RichText::new(format!("{hour:02}:00"))
                                    .size(16.0)
                                    .strong()
                                    .color(Color32::from_rgb(109, 194, 207)),
                            ),
                        );
                        for (idx, (name, _)) in RESOURCES.iter().enumerate() {
                            let booking = rs.iter().find(|r| {
                                r.resource == *name && r.start == hour && r.status != "cancelada"
                            });
                            let (fill, text) = if let Some(r) = booking {
                                (
                                    if r.status == "completada" {
                                        Color32::from_rgb(28, 58, 58)
                                    } else if idx >= 6 {
                                        Color32::from_rgb(81, 64, 34)
                                    } else {
                                        Color32::from_rgb(24, 63, 77)
                                    },
                                    format!(
                                        "{}\n{}",
                                        r.customer,
                                        if r.status == "completada" && !r.games.is_empty() {
                                            r.games.as_str()
                                        } else {
                                            "Reservado"
                                        }
                                    ),
                                )
                            } else {
                                (Color32::from_rgb(17, 30, 47), "LIBRE".to_string())
                            };
                            egui::Frame::new()
                                .fill(fill)
                                .stroke(egui::Stroke::new(
                                    1.0_f32,
                                    if idx >= 6 {
                                        Color32::from_rgb(117, 92, 47)
                                    } else {
                                        Color32::from_rgb(35, 58, 78)
                                    },
                                ))
                                .corner_radius(egui::CornerRadius::same(9))
                                .inner_margin(egui::Margin::symmetric(7, 8))
                                .show(ui, |ui| {
                                    ui.set_min_size(Vec2::new(92.0, 46.0));
                                    ui.vertical_centered(|ui| {
                                        if let Some(r) = booking {
                                            let customer =
                                                r.customer.chars().take(13).collect::<String>();
                                            ui.label(
                                                RichText::new(customer)
                                                    .strong()
                                                    .color(Color32::WHITE),
                                            );
                                            ui.label(
                                                RichText::new(if r.status == "completada" {
                                                    r.games.as_str()
                                                } else {
                                                    "Reservado"
                                                })
                                                .small()
                                                .color(Color32::from_rgb(150, 202, 207)),
                                            );
                                        } else {
                                            ui.label(
                                                RichText::new(text)
                                                    .small()
                                                    .color(Color32::from_rgb(95, 132, 151)),
                                            );
                                        }
                                    });
                                });
                        }
                        ui.end_row();
                    }
                });
        });
        ui.add_space(8.0);
        ui.label(RichText::new("Dos simuladores Oro  ·  seis simuladores Plata  ·  cada reserva ocupa un bloque de 60 minutos.").small().color(Color32::from_rgb(132, 153, 175)));
    }
    fn ui_reservations(&mut self, ui: &mut egui::Ui) {
        ui.label(
            RichText::new("RESERVAS  /  ACCESO RÁPIDO")
                .small()
                .color(Color32::from_rgb(91, 186, 200)),
        );
        ui.columns(2, |columns| {
            columns[0].vertical(|ui| {
        ui.heading(
            RichText::new(
                if self.editing_group_id > 0 || self.editing_reservation_id > 0 {
                    format!(
                        "Modificar reserva #{}",
                        if self.editing_group_id > 0 {
                            self.editing_group_id
                        } else {
                            self.editing_reservation_id
                        }
                    )
                } else {
                    "Nueva reserva grupal".into()
                },
            )
            .size(27.0)
            .strong(),
        );
        ui.label(
            RichText::new(if self.editing_reservation_id > 0 {
                "Reserva individual heredada: elegí una estación, una hora y revisá el nuevo total."
            } else {
                "Sumá todas las estaciones que necesita el cliente. Se reservarán juntas en el mismo horario."
            })
                .color(Color32::from_rgb(151, 171, 192)),
        );
        ui.add_space(10.0);
        ui.label(
            RichText::new("NOMBRE DEL CLIENTE")
                .small()
                .color(Color32::from_rgb(111, 141, 164)),
        );
        ui.add_sized(
            [ui.available_width().min(430.0), 40.0],
            egui::TextEdit::singleline(&mut self.customer)
                .hint_text("¿A nombre de quién reservamos?"),
        );
        ui.add_space(14.0);
        ui.horizontal(|ui| {
            ui.heading(
                RichText::new("01")
                    .color(Color32::from_rgb(91, 186, 200))
                    .size(18.0),
            );
            ui.heading("Elegí las estaciones");
            ui.label(
                RichText::new(format!("{} seleccionadas", self.selected_resources.len()))
                    .color(Color32::from_rgb(230, 190, 96)),
            );
        });
        ui.horizontal_wrapped(|ui| {
            ui.label(
                RichText::new("SELECCIÓN RÁPIDA")
                    .small()
                    .color(Color32::from_rgb(111, 141, 164)),
            );
            for (label, first, last) in [
                ("Plata ×6", 0, 6),
                ("Oro ×2", 6, 8),
                ("PS5 ×2", 8, 10),
                ("VR ×2", 10, 12),
            ] {
                let all_selected =
                    (first..last).all(|index| self.selected_resources.contains(&index));
                if ui.selectable_label(all_selected, label).clicked() {
                    if all_selected {
                        for index in first..last {
                            self.selected_resources.remove(&index);
                        }
                    } else {
                        for index in first..last {
                            self.selected_resources.insert(index);
                        }
                    }
                }
            }
            if ui.small_button("Limpiar").clicked() {
                self.selected_resources.clear();
            }
        });
        let reservations = self.reservations(&self.day);
        egui::Grid::new("resource_buttons")
            .num_columns(4)
            .spacing(Vec2::new(10.0, 10.0))
            .show(ui, |ui| {
                for (i, (name, _)) in RESOURCES.iter().enumerate() {
                    let selected = self.selected_resources.contains(&i);
                    let gold = i == 6 || i == 7;
                    let tint = if gold {
                        Color32::from_rgb(230, 188, 94)
                    } else {
                        Color32::from_rgb(107, 194, 209)
                    };
                    let fill = if selected {
                        if gold {
                            Color32::from_rgb(80, 61, 31)
                        } else {
                            Color32::from_rgb(24, 66, 80)
                        }
                    } else {
                        Color32::from_rgb(20, 31, 48)
                    };
                    let available = !reservations.iter().any(|r| {
                        r.resource == *name
                            && r.start == self.start_hour as i64
                            && r.status != "cancelada"
                            && (self.editing_group_id <= 0 || r.group_id != self.editing_group_id)
                            && (self.editing_reservation_id <= 0
                                || r.id != self.editing_reservation_id)
                    });
                    let tier = if gold {
                        "ORO"
                    } else if i < 6 {
                        "PLATA"
                    } else if i < 10 {
                        "PS5"
                    } else {
                        "VR"
                    };
                    let state = if selected && available {
                        "AGREGADA · LIBRE".to_string()
                    } else if selected {
                        format!("AGREGADA · ELEGÍ OTRA HORA")
                    } else if available {
                        "DISPONIBLE".to_string()
                    } else {
                        format!("OCUPADA A LAS {:02}:00", self.start_hour)
                    };
                    let label = format!("{}\n{}\n{}", tier, name, state);
                    let response = ui.add_sized(
                        [150.0, 64.0],
                        egui::Button::new(
                            RichText::new(label)
                                .color(if selected { Color32::WHITE } else { tint })
                                .strong(),
                        )
                        .fill(fill)
                        .stroke(egui::Stroke::new(
                            if selected { 1.6_f32 } else { 0.8_f32 },
                            if selected {
                                tint
                            } else {
                                Color32::from_rgb(47, 65, 86)
                            },
                        )),
                    );
                    if response.clicked() {
                        if self.editing_reservation_id > 0 {
                            self.selected_resources.clear();
                            self.selected_resources.insert(i);
                        } else if selected {
                            self.selected_resources.remove(&i);
                        } else {
                            self.selected_resources.insert(i);
                        }
                    }
                    if (i + 1) % 4 == 0 {
                        ui.end_row();
                    }
                }
            });
        ui.add_space(12.0);
        ui.horizontal(|ui| {
            ui.heading(
                RichText::new("02")
                    .color(Color32::from_rgb(91, 186, 200))
                    .size(18.0),
            );
            ui.heading("Elegí la hora");
        });
        ui.horizontal_wrapped(|ui| {
            for hour in 16..24 {
                let occupied = self.selected_resources.iter().any(|index| {
                    reservations.iter().any(|r| {
                        r.resource == RESOURCES[*index].0
                            && r.start == hour
                            && r.status != "cancelada"
                            && (self.editing_group_id <= 0 || r.group_id != self.editing_group_id)
                            && (self.editing_reservation_id <= 0
                                || r.id != self.editing_reservation_id)
                    })
                });
                let selected = self.start_hour == hour as i32;
                let color = if occupied {
                    Color32::from_rgb(90, 104, 117)
                } else if selected {
                    Color32::from_rgb(30, 138, 157)
                } else {
                    Color32::from_rgb(23, 38, 57)
                };
                if ui
                    .add_enabled(
                        !occupied,
                        egui::Button::new(
                            RichText::new(format!("{hour:02}:00")).size(16.0).strong(),
                        )
                        .min_size(Vec2::new(96.0, 48.0))
                        .fill(color),
                    )
                    .clicked()
                {
                    self.start_hour = hour as i32;
                }
            }
        });
        let (quoted_stations, promo_name, discount) =
            self.reservation_quote(&self.selected_resources, self.start_hour);
        let price = quoted_stations
            .iter()
            .fold(0_i64, |sum, (_, amount, _)| sum.saturating_add(*amount));
        let station_names = self
            .selected_resources
            .iter()
            .map(|index| RESOURCES[*index].0)
            .collect::<Vec<_>>()
            .join(" · ");
        ui.add_space(14.0);
        egui::Frame::new()
            .fill(Color32::from_rgb(19, 32, 49))
            .corner_radius(egui::CornerRadius::same(12))
            .inner_margin(egui::Margin::symmetric(14, 12))
            .show(ui, |ui| {
                ui.horizontal(|ui| {
                    ui.vertical(|ui| {
                        ui.label(
                            RichText::new(format!(
                                "{} estación(es)  ·  {:02}:00–{:02}:00",
                                self.selected_resources.len(),
                                self.start_hour,
                                self.start_hour + 1
                            ))
                            .strong()
                            .size(16.0),
                        );
                        ui.label(
                            RichText::new(if station_names.is_empty() {
                                "Seleccioná las estaciones que querés reservar".to_string()
                            } else {
                                station_names.clone()
                            })
                            .small()
                            .color(Color32::from_rgb(152, 185, 199)),
                        );
                        ui.label(
                            RichText::new(if promo_name.is_empty() {
                                "Bloque fijo de 60 minutos · precio por estación".to_string()
                            } else {
                                format!("Oferta {} · ahorrás {}", promo_name, money(discount))
                            })
                            .small()
                            .color(if promo_name.is_empty() {
                                Color32::GRAY
                            } else {
                                Color32::from_rgb(124, 214, 177)
                            }),
                        );
                    });
                    ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                        ui.strong(
                            RichText::new(money(price))
                                .size(21.0)
                                .color(Color32::from_rgb(232, 194, 107)),
                        );
                    });
                });
            });
        ui.add_space(10.0);
        if ui
            .add_enabled(
                !self.customer.trim().is_empty()
                    && !self.selected_resources.is_empty()
                    && !self.selected_resources.iter().any(|index| {
                        reservations.iter().any(|r| {
                            r.resource == RESOURCES[*index].0
                                && r.start == self.start_hour as i64
                                && r.status != "cancelada"
                                && (self.editing_group_id <= 0
                                    || r.group_id != self.editing_group_id)
                                && (self.editing_reservation_id <= 0
                                    || r.id != self.editing_reservation_id)
                        })
                    }),
                egui::Button::new(
                    RichText::new(
                        if self.editing_group_id > 0 || self.editing_reservation_id > 0 {
                            "GUARDAR CAMBIOS"
                        } else {
                            "CONFIRMAR RESERVA"
                        },
                    )
                    .strong(),
                )
                .min_size(Vec2::new(260.0, 46.0))
                .fill(Color32::from_rgb(29, 120, 137)),
            )
            .clicked()
        {
            self.create_reservation();
        }
        if (self.editing_group_id > 0 || self.editing_reservation_id > 0)
            && ui.button("Salir de la edición").clicked()
        {
            self.editing_group_id = 0;
            self.editing_reservation_id = 0;
            self.customer.clear();
            self.selected_resources.clear();
            self.message = "Edición descartada.".into();
        }
            });
            columns[1].vertical(|ui| {
                ui.heading(format!("Turnos de hoy · {}", self.day));
                self.reservation_list(ui);
            });
        });
    }
    fn reservation_list(&mut self, ui: &mut egui::Ui) {
        let rows = self.reservations(&self.day);
        let mut groups: BTreeMap<i64, Vec<Reservation>> = BTreeMap::new();
        for row in rows {
            let key = if row.group_id > 0 {
                row.group_id
            } else {
                -row.id
            };
            groups.entry(key).or_default().push(row);
        }
        for (_, stations) in groups {
            let first = &stations[0];
            let group_id = first.group_id;
            let reservation_id = if group_id == 0 { first.id } else { 0 };
            let total = stations
                .iter()
                .fold(0_i64, |sum, s| sum.saturating_add(s.total));
            let paid = self.paid_for(group_id, reservation_id);
            let canceled = stations.iter().all(|r| r.status == "cancelada");
            let balance = if canceled {
                0
            } else {
                total.saturating_sub(paid)
            };
            let editable = stations.iter().all(|r| r.status == "reservada");
            let cancellable = stations.iter().all(|r| r.status == "reservada");
            let payable = stations
                .iter()
                .all(|r| r.status == "reservada" || r.status == "completada");
            let mut edit = false;
            let mut cancel = false;
            let mut payment_action = None;
            egui::Frame::new()
                .fill(Color32::from_rgb(18, 30, 46))
                .corner_radius(egui::CornerRadius::same(10))
                .inner_margin(egui::Margin::symmetric(12, 6))
                .show(ui, |ui| {
                    ui.horizontal(|ui| {
                                ui.strong(format!("{:02}:00", first.start));
                                ui.strong(&first.customer);
                                ui.label(format!("{} estación(es)", stations.len()));
                                if first.group_id > 0 {
                                    let promo = self
                                        .db
                                        .query_row(
                                            "SELECT promo_name,discount_cents FROM reservation_groups WHERE id=?1",
                                            [first.group_id],
                                            |r| Ok((r.get::<_, String>(0)?, r.get::<_, i64>(1)?)),
                                        )
                                        .unwrap_or_default();
                                    if !promo.0.is_empty() {
                                        ui.label(
                                            RichText::new(format!("OFERTA: {} · ahorro {}", promo.0, money(promo.1)))
                                                .small()
                                                .color(Color32::from_rgb(124, 214, 177)),
                                        );
                                    }
                                }
                                ui.with_layout(
                                    egui::Layout::right_to_left(egui::Align::Center),
                                    |ui| {
                                        ui.strong(
                                            RichText::new(money(
                                                stations.iter().map(|s| s.total).sum(),
                                            ))
                                            .color(Color32::from_rgb(230, 190, 96)),
                                        );
                                    },
                                );
                    });
                    ui.collapsing("Estaciones, pagos y acciones", |ui| {
                            ui.horizontal_wrapped(|ui| {
                                for station in &stations {
                                    let label = if station.status == "completada"
                                        && !station.games.is_empty()
                                    {
                                        format!("{} · {}", station.resource, station.games)
                                    } else {
                                        format!("{} · {}", station.resource, station.status)
                                    };
                                    ui.label(
                                        RichText::new(label)
                                            .small()
                                            .color(Color32::from_rgb(148, 179, 195)),
                                    );
                                }
                            });
                            ui.add_space(4.0);
                            ui.horizontal_wrapped(|ui| {
                                if canceled {
                                    ui.label(RichText::new("RESERVA CANCELADA").color(Color32::from_rgb(224, 129, 119)).strong());
                                } else {
                                    ui.label(format!("Pagado {} · Saldo {}", money(paid), money(balance)));
                                }
                                if !canceled {
                                    if balance == 0 {
                                        ui.label(RichText::new("PAGADA").color(Color32::from_rgb(124, 214, 177)).strong());
                                    } else if paid > 0 {
                                        ui.label(RichText::new("SEÑA / PAGO PARCIAL").color(Color32::from_rgb(230, 190, 96)));
                                    } else {
                                        ui.label(RichText::new("SIN PAGOS").color(Color32::from_rgb(148, 179, 195)));
                                    }
                                }
                            });
                            if cancellable && editable {
                                ui.horizontal_wrapped(|ui| {
                                    if editable && ui.button("Modificar").clicked() { edit = true; }
                                    if ui.button(if paid > 0 { "Devolver y cancelar" } else { "Cancelar reserva" }).clicked() { cancel = true; }
                                });
                            }
                            if payable {
                                ui.horizontal_wrapped(|ui| {
                                    ui.label(if cancellable && paid > 0 { "Cobro o devolución por" } else { "Recibir por" });
                                    ui.selectable_value(&mut self.reservation_payment_transfer, false, "Efectivo");
                                    ui.selectable_value(&mut self.reservation_payment_transfer, true, "Transferencia");
                                    let deposit_target = (total + 1) / 2;
                                    if paid < deposit_target && ui.button(format!("Registrar seña 50% · {}", money(deposit_target - paid))).clicked() {
                                        payment_action = Some(reservations::PaymentAction::Deposit);
                                    }
                                    if balance > 0 && ui.button(format!("Cobrar saldo · {}", money(balance))).clicked() {
                                        payment_action = Some(reservations::PaymentAction::Balance);
                                    }
                                });
                                ui.horizontal(|ui| {
                                    ui.label("Otro pago parcial ($)");
                                    ui.add_sized([110.0, 28.0], egui::TextEdit::singleline(&mut self.reservation_partial_amount).hint_text("Monto"));
                                    if balance > 0 && ui.button("Registrar pago").clicked() {
                                        payment_action = Some(reservations::PaymentAction::Partial);
                                    }
                                });
                            }
                    });
                });
            if edit {
                self.begin_edit_group(group_id, &stations);
            } else if cancel {
                self.pending_cancel = Some((group_id, reservation_id));
            } else if let Some(action) = payment_action {
                self.record_reservation_payment(
                    group_id,
                    reservation_id,
                    action,
                    pesos_input(&self.reservation_partial_amount),
                );
                self.reservation_partial_amount.clear();
            }
        }
    }
    fn ui_sales(&mut self, ui: &mut egui::Ui) {
        ui.label(
            RichText::new("PUNTO DE VENTA  /  BEBIDAS")
                .small()
                .color(Color32::from_rgb(91, 186, 200)),
        );
        ui.heading(RichText::new("Armar una compra").size(27.0).strong());
        ui.horizontal(|ui| {
            ui.selectable_value(
                &mut self.pay_transfer,
                false,
                RichText::new("EFECTIVO").strong(),
            );
            ui.selectable_value(
                &mut self.pay_transfer,
                true,
                RichText::new("TRANSFERENCIA").strong(),
            );
        });
        ui.label(
            RichText::new("Tocá un producto para sumarlo al carrito")
                .color(Color32::from_rgb(150, 171, 193)),
        );
        ui.add_space(8.0);
        egui::Grid::new("product_tiles")
            .num_columns(4)
            .spacing(Vec2::new(12.0, 12.0))
            .show(ui, |ui| {
                for (i, p) in self.products().into_iter().enumerate() {
                    let in_cart = *self.cart.get(&p.id).unwrap_or(&0);
                    let response = ui.add_enabled(
                        p.stock > in_cart,
                        egui::Button::new(
                            RichText::new(format!(
                                "BEBIDA\n{}\n{}\nStock {}",
                                p.name,
                                money(p.price),
                                p.stock
                            ))
                            .strong()
                            .size(14.0)
                            .color(Color32::from_rgb(222, 233, 240)),
                        )
                        .min_size(Vec2::new(142.0, 126.0))
                        .fill(Color32::from_rgb(22, 39, 57))
                        .stroke(egui::Stroke::new(1.0_f32, Color32::from_rgb(43, 70, 89))),
                    );
                    if response.clicked() {
                        self.add_to_cart(p.id);
                    }
                    if (i + 1) % 4 == 0 {
                        ui.end_row();
                    }
                }
            });
        ui.add_space(12.0);
        egui::Frame::new()
            .fill(Color32::from_rgb(18, 29, 45))
            .corner_radius(egui::CornerRadius::same(13))
            .inner_margin(egui::Margin::same(14))
            .show(ui, |ui| {
                let products = self.products();
                ui.horizontal(|ui| {
                    ui.heading("Carrito");
                    ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                        let total = self
                            .cart
                            .iter()
                            .filter_map(|(id, q)| {
                                products.iter().find(|p| p.id == *id).map(|p| p.price * *q)
                            })
                            .sum::<i64>();
                        ui.strong(
                            RichText::new(money(total))
                                .size(21.0)
                                .color(Color32::from_rgb(231, 192, 105)),
                        );
                    });
                });
                let mut remove = Vec::new();
                for (id, quantity) in &mut self.cart {
                    if let Some(p) = products.iter().find(|p| p.id == *id) {
                        ui.horizontal(|ui| {
                            ui.label(&p.name);
                            ui.label(format!("{} × {}", quantity, money(p.price)));
                            if ui.small_button("−").clicked() {
                                *quantity -= 1;
                            }
                            if ui.small_button("+").clicked() && *quantity < p.stock {
                                *quantity += 1;
                            }
                            if *quantity <= 0 {
                                remove.push(*id);
                            }
                        });
                    }
                }
                for id in remove {
                    self.cart.remove(&id);
                }
                if self.cart.is_empty() {
                    ui.label(RichText::new("Todavía no agregaste productos.").color(Color32::GRAY));
                }
                if ui
                    .add_enabled(
                        !self.cart.is_empty(),
                        egui::Button::new(RichText::new("COBRAR COMPRA").strong())
                            .min_size(Vec2::new(210.0, 44.0))
                            .fill(Color32::from_rgb(28, 125, 139)),
                    )
                    .clicked()
                {
                    self.checkout_cart();
                }
            });
        ui.add_space(16.0);
        ui.collapsing("Cerrar una sesión reservada", |ui| {
            ui.label("El pago y la seña se registran por separado en Reservas. Esta acción solo marca el uso y el juego de la estación.");
            ui.label("Juego asignado a esta franja");
            ui.horizontal_wrapped(|ui| {
                for game in std::iter::once("Sin clasificar").chain(GAME_OPTIONS) {
                    let selected = self.session_game == game;
                    if ui.selectable_label(selected, game).clicked() {
                        self.session_game = game.into();
                    }
                }
            });
            let rows = self.reservations(&self.day);
            for r in rows.into_iter().filter(|r| r.status == "reservada") {
                ui.horizontal(|ui| {
                    ui.label(format!(
                        "{} · {} · {:02}:00 · {}",
                        r.resource,
                        r.customer,
                        r.start,
                        money(r.total)
                    ));
                    if ui.button("Completar sesión").clicked() {
                        self.register_session(&r);
                    }
                });
            }
        });
    }
    fn ui_products(&mut self, ui: &mut egui::Ui) {
        ui.label(
            RichText::new("CATÁLOGO  /  INVENTARIO")
                .small()
                .color(Color32::from_rgb(91, 186, 200)),
        );
        ui.heading(RichText::new("Bebidas y stock").size(27.0).strong());
        ui.label(
            "El recargo porcentual se suma al costo de compra para sugerir el precio de venta.",
        );
        ui.horizontal(|ui| {
            if ui.button("↓  Descargar plantilla Excel").clicked() {
                self.export_product_template();
            }
            if ui.button("↑  Importar productos Excel").clicked() {
                self.import_products();
            }
        });
        ui.separator();
        ui.heading("Agregar producto o reponer stock");
        egui::Grid::new("product_form")
            .num_columns(2)
            .spacing(Vec2::new(12.0, 8.0))
            .show(ui, |ui| {
                ui.label("Nombre");
                ui.add_sized(
                    [260.0, 32.0],
                    egui::TextEdit::singleline(&mut self.product_name)
                        .hint_text("Ej. Agua sin gas"),
                );
                ui.end_row();
                ui.label("Costo unitario ($)");
                ui.add_sized(
                    [260.0, 32.0],
                    egui::TextEdit::singleline(&mut self.product_cost).hint_text("Costo de compra"),
                );
                ui.end_row();
                ui.label("Recargo (%)");
                ui.add_sized(
                    [120.0, 32.0],
                    egui::TextEdit::singleline(&mut self.product_margin),
                );
                ui.end_row();
                ui.label("Unidades compradas");
                ui.add_sized(
                    [120.0, 32.0],
                    egui::TextEdit::singleline(&mut self.product_stock),
                );
                ui.end_row();
            });
        let computed_sale = sale_price(
            pesos_input(&self.product_cost),
            margin_basis_points(&self.product_margin),
        );
        ui.horizontal(|ui| {
            ui.label("Precio sugerido");
            ui.strong(RichText::new(money(computed_sale)).color(Color32::from_rgb(231, 192, 105)));
            if ui.button("Guardar / reponer").clicked() {
                self.add_product();
            }
        });
        ui.separator();
        ui.heading("Catálogo actual");
        egui::Grid::new("product_list")
            .striped(true)
            .num_columns(5)
            .show(ui, |ui| {
                for header in ["Producto", "Costo", "Recargo", "Venta", "Stock"] {
                    ui.strong(header);
                }
                ui.end_row();
                for p in self.products() {
                    ui.label(p.name);
                    ui.label(money(p.cost));
                    ui.label(format!("{:.2}%", p.margin_bps as f64 / 100.0));
                    ui.strong(money(p.price));
                    ui.label(p.stock.to_string());
                    ui.end_row();
                }
            });
    }
    fn ui_prices(&mut self, ui: &mut egui::Ui) {
        ui.heading("Tarifas base por hora");
        ui.label("Se guardan como pesos enteros. El valor se copia a nuevas reservas; las existentes mantienen su importe histórico.");
        for (label, value) in [
            ("Simuladores Plata", &mut self.price_sim),
            ("Simuladores Oro", &mut self.price_gold),
            ("PS5", &mut self.price_play),
            ("VR", &mut self.price_vr),
        ] {
            ui.horizontal(|ui| {
                ui.label(label);
                ui.add_sized([130.0, 26.0], egui::TextEdit::singleline(value));
                ui.label("ARS / hora");
            });
        }
        if ui.button("Guardar tarifas").clicked() {
            self.save_prices();
        }
        ui.label("Las nuevas reservas toman la tarifa vigente y conservan ese importe. Configurá ofertas por horario en Promociones.");
    }
    fn ui_promotions(&mut self, ui: &mut egui::Ui) {
        ui.label(
            RichText::new("TARIFAS / REGLAS COMERCIALES")
                .small()
                .color(Color32::from_rgb(91, 186, 200)),
        );
        ui.heading(RichText::new("Promociones automáticas").size(27.0).strong());
        ui.label("La app elige una sola oferta aplicable, la más conveniente, y descuenta las estaciones de menor valor de cada grupo completo. Horario de fin incluido.");
        ui.add_space(10.0);
        ui.label("Nombre de la oferta");
        ui.add_sized(
            [320.0, 34.0],
            egui::TextEdit::singleline(&mut self.promo_name).hint_text("Ej.: Promo tarde"),
        );
        ui.horizontal_wrapped(|ui| {
            ui.label("Mecánica");
            for (label, buy, pay) in [
                ("2 por 1", "2", "1"),
                ("3 por 2", "3", "2"),
                ("4 por 3", "4", "3"),
            ] {
                if ui
                    .selectable_label(self.promo_buy == buy && self.promo_pay == pay, label)
                    .clicked()
                {
                    self.promo_buy = buy.into();
                    self.promo_pay = pay.into();
                }
            }
        });
        ui.horizontal_wrapped(|ui| {
            ui.label("Aplicar a");
            for (label, kind) in [
                ("Todos", "all"),
                ("Plata", "sim_silver"),
                ("Oro", "sim_gold"),
                ("PS5", "play"),
                ("VR", "vr"),
            ] {
                ui.selectable_value(&mut self.promo_kind, kind.to_string(), label);
            }
        });
        ui.label("Vigencia por hora de inicio de la reserva (ambos extremos incluidos)");
        ui.horizontal_wrapped(|ui| {
            ui.label("Desde");
            for hour in 16..=23 {
                if ui
                    .selectable_label(self.promo_start == hour, format!("{hour:02}:00"))
                    .clicked()
                {
                    self.promo_start = hour;
                    if self.promo_end < hour {
                        self.promo_end = hour;
                    }
                }
            }
        });
        ui.horizontal_wrapped(|ui| {
            ui.label("Hasta");
            for hour in self.promo_start..=23 {
                if ui
                    .selectable_label(self.promo_end == hour, format!("{hour:02}:00"))
                    .clicked()
                {
                    self.promo_end = hour;
                }
            }
        });
        if ui.button("Guardar promoción").clicked() {
            self.add_promotion();
        }
        ui.separator();
        ui.heading("Promociones activas");
        for promo in self.promotions() {
            ui.horizontal(|ui| {
                ui.label(
                    RichText::new(format!(
                        "{} por {} · {} · {:02}:00–{:02}:00 · {}",
                        promo.buy_qty,
                        promo.pay_qty,
                        promo.name,
                        promo.start_hour,
                        promo.end_hour,
                        match promo.resource_kind.as_str() {
                            "all" => "Todas las estaciones",
                            "sim_silver" => "Simuladores Plata",
                            "sim_gold" => "Simuladores Oro",
                            "play" => "PS5",
                            "vr" => "VR",
                            _ => "Categoría",
                        }
                    ))
                    .strong(),
                );
                if ui.small_button("Desactivar").clicked() {
                    match self
                        .db
                        .execute("UPDATE promotions SET active=0 WHERE id=?1", [promo.id])
                    {
                        Ok(_) => self.message = format!("Promoción '{}' desactivada.", promo.name),
                        Err(e) => self.message = format!("No se pudo desactivar: {e}"),
                    }
                }
            });
        }
        ui.label(RichText::new("Ejemplo: 2 por 1 · Todas · desde 16:00 hasta 19:00 incluye reservas que comienzan a las 16, 17, 18 y 19.").small().color(Color32::from_rgb(148, 179, 195)));
    }
    fn ui_timers(&mut self, ui: &mut egui::Ui) {
        ui.ctx().request_repaint_after(Duration::from_secs(1));
        ui.label(
            RichText::new("SEGUIMIENTO EN TIEMPO REAL")
                .small()
                .color(Color32::from_rgb(91, 186, 200)),
        );
        ui.heading(RichText::new("Cronómetros de juego").size(27.0).strong());
        ui.label("Iniciá un contador para cada PS5 o puesto de realidad virtual. El aro muestra el tiempo restante y sigue corriendo aunque cierres y abras VRBox.");
        ui.add_space(12.0);
        ui.horizontal_wrapped(|ui| {
            ui.label("Puesto");
            for (name, _) in RESOURCES
                .iter()
                .filter(|(_, kind)| *kind == "play" || *kind == "vr")
            {
                ui.selectable_value(&mut self.timer_resource, (*name).to_string(), *name);
            }
        });
        ui.horizontal(|ui| {
            ui.label("Cliente (opcional)");
            ui.add_sized(
                [220.0, 32.0],
                egui::TextEdit::singleline(&mut self.timer_customer).hint_text("Nombre"),
            );
            ui.label("Minutos");
            ui.add_sized(
                [80.0, 32.0],
                egui::TextEdit::singleline(&mut self.timer_minutes),
            );
            if ui.button("Iniciar cronómetro").clicked() {
                self.add_timer();
            }
        });
        ui.separator();
        let timers = self.game_timers();
        if timers.is_empty() {
            ui.add_space(20.0);
            ui.label(RichText::new("Todavía no hay cronómetros activos. Elegí un puesto y una duración para comenzar.").color(Color32::from_rgb(148, 179, 195)));
        }
        egui::Grid::new("timer_cards")
            .num_columns(2)
            .spacing(Vec2::new(12.0, 12.0))
            .show(ui, |ui| {
                for (index, timer) in timers.iter().enumerate() {
                    let remaining = timer.ends_at.saturating_sub(Local::now().timestamp());
                    egui::Frame::new()
                        .fill(Color32::from_rgb(18, 30, 46))
                        .corner_radius(egui::CornerRadius::same(12))
                        .inner_margin(egui::Margin::symmetric(12, 10))
                        .show(ui, |ui| {
                            ui.horizontal(|ui| {
                                timer_ring(ui, remaining, timer.total_secs);
                                ui.vertical(|ui| {
                                    ui.heading(&timer.resource);
                                    ui.label(if timer.customer.is_empty() {
                                        "Sin nombre de cliente"
                                    } else {
                                        &timer.customer
                                    });
                                    let started = chrono::DateTime::<chrono::Utc>::from_timestamp(
                                        timer.started_at,
                                        0,
                                    )
                                    .map(|d| {
                                        d.with_timezone(&Local)
                                            .format("Iniciado a las %H:%M")
                                            .to_string()
                                    })
                                    .unwrap_or_default();
                                    ui.label(
                                        RichText::new(started)
                                            .small()
                                            .color(Color32::from_rgb(148, 179, 195)),
                                    );
                                    ui.horizontal(|ui| {
                                        if ui.small_button("- 5 min").clicked() {
                                            self.adjust_timer(timer.id, -5);
                                        }
                                        if ui.small_button("+ 5 min").clicked() {
                                            self.adjust_timer(timer.id, 5);
                                        }
                                        if ui.small_button("Eliminar").clicked() {
                                            self.delete_timer(timer.id);
                                        }
                                    });
                                });
                            });
                        });
                    if index % 2 == 1 {
                        ui.end_row();
                    }
                }
            });
        ui.add_space(8.0);
        ui.label(RichText::new("Los botones de ajuste modifican el vencimiento en 5 minutos. Los cronómetros finalizados quedan visibles hasta que se eliminen; sumar tiempo puede reactivarlos.").small().color(Color32::from_rgb(148, 179, 195)));
    }
    fn ui_closing(&mut self, ui: &mut egui::Ui) {
        ui.heading(format!("Cierre de caja · {}", self.day));
        let (cash, transfer, drinks, pending) = self.totals();
        ui.horizontal_wrapped(|ui| {
            stat(ui, "Efectivo esperado", money(cash));
            stat(ui, "Transferencias esperadas", money(transfer));
            stat(ui, "Ventas bebidas", money(drinks));
            stat(ui, "Reservas pendientes", money(pending));
        });
        ui.add_space(8.0);
        ui.label("Contá el dinero físico e informá por separado lo acreditado por transferencia.");
        ui.horizontal(|ui| {
            ui.label("Efectivo contado $");
            ui.text_edit_singleline(&mut self.counted_cash);
            ui.label("Transferencias $");
            ui.text_edit_singleline(&mut self.counted_transfer);
        });
        ui.horizontal(|ui| {
            ui.label("Observaciones");
            ui.text_edit_singleline(&mut self.closing_note);
        });
        ui.label(format!(
            "Diferencia efectivo: {} · transferencia: {}",
            money(pesos_input(&self.counted_cash) - cash),
            money(pesos_input(&self.counted_transfer) - transfer)
        ));
        if ui
            .button("Guardar cierre y exportar Excel (.xlsx)")
            .clicked()
        {
            self.close_and_export();
        }
        ui.small(
            "Un cierre interno no sustituye los comprobantes fiscales ni la facturación exigible.",
        );
    }
    fn ui_events(&mut self, ui: &mut egui::Ui) {
        ui.heading("Cumpleaños, eventos y alertas");
        ui.horizontal(|ui| {
            ui.label("Fecha YYYY-MM-DD");
            ui.text_edit_singleline(&mut self.event_day);
            egui::ComboBox::from_id_salt("event_kind")
                .selected_text(&self.event_kind)
                .show_ui(ui, |ui| {
                    for k in ["Cumpleaños", "Torneo", "Evento", "Otro"] {
                        ui.selectable_value(&mut self.event_kind, k.into(), k);
                    }
                });
        });
        ui.horizontal(|ui| {
            ui.label("Título");
            ui.text_edit_singleline(&mut self.event_title);
            ui.label("Cliente");
            ui.text_edit_singleline(&mut self.event_customer);
            if ui.button("Agregar").clicked() {
                self.add_event();
            }
        });
        ui.separator();
        for e in self.events() {
            ui.label(format!(
                "{} · {} · {} · {}",
                e.day, e.kind, e.title, e.customer
            ));
        }
        ui.label("Las alertas usan la fecha operativa visible arriba; recordatorios recurrentes de cumpleaños requieren una fecha de nacimiento anual.");
    }
}
fn stat(ui: &mut egui::Ui, label: &str, value: String) {
    egui::Frame::group(ui.style()).show(ui, |ui| {
        ui.vertical(|ui| {
            ui.small(label);
            ui.strong(value);
        });
    });
}

fn main() -> eframe::Result<()> {
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_inner_size(Vec2::new(1320.0, 820.0))
            .with_min_inner_size(Vec2::new(960.0, 640.0)),
        ..Default::default()
    };
    eframe::run_native(
        "VRBox · Control de ciber",
        options,
        Box::new(|_cc| Ok(Box::new(VrBoxApp::new()))),
    )
}
