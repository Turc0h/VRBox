use calamine::{Reader, open_workbook_auto};
use chrono::{Local, NaiveDate};
use eframe::egui::{self, Color32, RichText, Vec2};
use rusqlite::{Connection, params};
use rust_xlsxwriter::{Format, Formula, Workbook};
use std::collections::BTreeMap;
use std::{path::PathBuf, time::Duration};

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
           status TEXT NOT NULL DEFAULT 'reservada', created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
           UNIQUE(service_date, resource, start_hour));
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

#[derive(PartialEq, Clone, Copy)]
enum Page {
    Agenda,
    Reservations,
    Sales,
    Products,
    Prices,
    Closing,
    Events,
}

struct VrBoxApp {
    db: Connection,
    page: Page,
    day: String,
    message: String,
    customer: String,
    resource_idx: usize,
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
            resource_idx: 0,
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
            .map(|v: i64| v * 100)
            .unwrap_or(0)
    }
    fn reservations(&self, day: &str) -> Vec<Reservation> {
        let mut stmt = match self.db.prepare("SELECT id,service_date,resource,customer,start_hour,duration,games,total_cents,status FROM reservations WHERE service_date=?1 ORDER BY start_hour,resource") { Ok(s) => s, Err(_) => return vec![] };
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
        if self.customer.trim().is_empty() {
            self.message = "Ingresá el nombre del cliente.".into();
            return;
        }
        if NaiveDate::parse_from_str(&self.day, "%Y-%m-%d").is_err() {
            self.message = "La fecha operativa debe tener formato AAAA-MM-DD.".into();
            return;
        }
        if !(16..=23).contains(&self.start_hour) {
            self.message = "El horario debe quedar dentro de la jornada de 16:00 a 00:00.".into();
            return;
        }
        let (resource, kind) = RESOURCES[self.resource_idx];
        let collision: i64 = self.db.query_row(
            "SELECT COUNT(*) FROM reservations WHERE service_date=?1 AND resource=?2 AND status!='cancelada' AND start_hour < ?3 AND start_hour + duration > ?4",
            params![self.day, resource, self.start_hour + 1, self.start_hour],
            |row| row.get(0),
        ).unwrap_or(1);
        if collision > 0 {
            self.message = "Ese puesto ya está ocupado en parte de ese horario.".into();
            return;
        }
        let games = String::new();
        let total = self.price_for_kind(kind);
        let result=self.db.execute("INSERT INTO reservations(service_date,resource,customer,start_hour,duration,games,total_cents,discount_cents,note) VALUES(?1,?2,?3,?4,1,?5,?6,0,'')",params![self.day,resource,self.customer.trim(),self.start_hour,games,total]);
        self.message = match result {
            Ok(_) => {
                self.customer.clear();
                "Reserva guardada.".into()
            }
            Err(e) => format!("No se pudo reservar: {e}"),
        };
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
        let tx = match self.db.transaction() {
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
            let result = tx.execute("INSERT INTO sales(category,item,quantity,unit_cents,total_cents,payment,service_date) VALUES('bebida',?1,?2,?3,?4,?5,?6)", params![p.name, quantity, p.price, total, payment, self.day])
                .and_then(|_| tx.execute("UPDATE products SET stock=stock-?1 WHERE id=?2 AND stock>=?1", params![quantity,p.id]));
            if let Err(e) = result {
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
        let tx = match self.db.transaction() {
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
            let _=self.db.execute("INSERT INTO settings(key,value) VALUES(?1,?2) ON CONFLICT(key) DO UPDATE SET value=excluded.value",params![key,value.parse::<i64>().unwrap_or(0).max(0).to_string()]);
        }
        self.message = "Tarifas guardadas. Se aplican a nuevas reservas.".into();
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
        let reserved=self.db.query_row("SELECT COALESCE(SUM(total_cents),0) FROM reservations WHERE service_date=?1 AND status='reservada'",[self.day.as_str()],|r|r.get(0)).unwrap_or(0);
        (row.0, row.1, row.2, reserved)
    }
    fn register_session(&mut self, r: &Reservation) {
        let payment = if self.pay_transfer {
            "transferencia"
        } else {
            "efectivo"
        };
        let game = self.session_game.clone();
        let tx = match self.db.transaction() {
            Ok(tx) => tx,
            Err(e) => {
                self.message = format!("No se pudo iniciar el cobro: {e}");
                return;
            }
        };
        let result = tx.execute("INSERT INTO sales(category,item,quantity,unit_cents,total_cents,payment,service_date) VALUES('sesion',?1,1,?2,?2,?3,?4)",params![format!("{} — {}",r.resource,game),r.total,payment,self.day])
            .and_then(|_| tx.execute("UPDATE reservations SET status='completada',games=?1 WHERE id=?2 AND status='reservada'",params![game,r.id]));
        if result.is_err() {
            self.message = "No se pudo completar el cobro de la sesión.".into();
            return;
        }
        match tx.commit() {
            Ok(()) => self.message = format!("Sesión cobrada · {} · {}.", game, money(r.total)),
            Err(e) => self.message = format!("No se pudo completar el cobro: {e}"),
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
            for col in 0..8 {
                sheet.set_column_width(col, 21).map_err(|e| e.to_string())?;
            }
        }
        {
            let sheet = workbook.add_worksheet();
            sheet.set_name("Por juego").map_err(|e| e.to_string())?;
            for (col, title) in [
                "Juego",
                "Efectivo centavos",
                "Transferencia centavos",
                "Total centavos",
            ]
            .iter()
            .enumerate()
            {
                sheet
                    .write_string(0, col as u16, *title)
                    .map_err(|e| e.to_string())?;
            }
            let mut totals: BTreeMap<String, (i64, i64)> = BTreeMap::new();
            let mut stmt = self.db.prepare("SELECT item,total_cents,payment FROM sales WHERE service_date=?1 AND category='sesion'").map_err(|e| e.to_string())?;
            let rows = stmt
                .query_map([self.day.as_str()], |r| {
                    Ok((
                        r.get::<_, String>(0)?,
                        r.get::<_, i64>(1)?,
                        r.get::<_, String>(2)?,
                    ))
                })
                .map_err(|e| e.to_string())?;
            for row in rows {
                let (item, total, payment) = row.map_err(|e| e.to_string())?;
                let games = item
                    .split(" — ")
                    .nth(1)
                    .unwrap_or("")
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
                    if payment == "efectivo" {
                        entry.0 += attributed;
                    } else {
                        entry.1 += attributed;
                    }
                }
            }
            sheet
                .write_string(
                    1,
                    0,
                    "Regla: sesiones multijuego distribuidas en partes iguales",
                )
                .map_err(|e| e.to_string())?;
            for (index, (game, (cash, transfer))) in totals.iter().enumerate() {
                let row = index as u32 + 2;
                sheet
                    .write_string(row, 0, game)
                    .map_err(|e| e.to_string())?;
                sheet
                    .write_number(row, 1, *cash as f64)
                    .map_err(|e| e.to_string())?;
                sheet
                    .write_number(row, 2, *transfer as f64)
                    .map_err(|e| e.to_string())?;
                sheet
                    .write_number(row, 3, (cash + transfer) as f64)
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
                .write_string(0, 1, "Horas reservadas cobradas")
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
                            ui.label(
                                RichText::new("FECHA OPERATIVA")
                                    .small()
                                    .color(Color32::GRAY),
                            );
                            ui.add_sized([130.0, 30.0], egui::TextEdit::singleline(&mut self.day));
                            ui.label(
                                RichText::new("●  SISTEMA LOCAL")
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
                    (Page::Agenda, "▦", "Agenda"),
                    (Page::Reservations, "＋", "Reservas"),
                    (Page::Sales, "▣", "Punto de venta"),
                    (Page::Products, "◈", "Bebidas y stock"),
                    (Page::Prices, "◇", "Tarifas"),
                    (Page::Closing, "◫", "Cierre de caja"),
                    (Page::Events, "✦", "Eventos y alertas"),
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
                                ui.label(RichText::new(icon).color(text_color).size(18.0));
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
                                    format!("◆ {name}")
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
        ui.label(RichText::new("◆ Dos simuladores Oro  ·  seis simuladores Plata  ·  cada reserva ocupa un bloque de 60 minutos.").small().color(Color32::from_rgb(132, 153, 175)));
    }
    fn ui_reservations(&mut self, ui: &mut egui::Ui) {
        ui.label(
            RichText::new("RESERVAS  /  ACCESO RÁPIDO")
                .small()
                .color(Color32::from_rgb(91, 186, 200)),
        );
        ui.heading(RichText::new("Nueva sesión").size(27.0).strong());
        ui.label(
            RichText::new("Elegí un equipo y una hora. Cada turno dura una hora.")
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
            ui.heading("Elegí tu espacio");
        });
        egui::Grid::new("resource_buttons")
            .num_columns(4)
            .spacing(Vec2::new(10.0, 10.0))
            .show(ui, |ui| {
                for (i, (name, _)) in RESOURCES.iter().enumerate() {
                    let selected = self.resource_idx == i;
                    let gold = i >= 6;
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
                    let label = format!("{}\n{}", if gold { "◆ ORO" } else { "◇ PLATA" }, name);
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
                        self.resource_idx = i;
                        let reservations = self.reservations(&self.day);
                        let blocked = |hour| {
                            reservations.iter().any(|r| {
                                r.resource == *name && r.start == hour && r.status != "cancelada"
                            })
                        };
                        if blocked(self.start_hour as i64) {
                            self.start_hour =
                                (16..24).find(|hour| !blocked(*hour)).unwrap_or(16) as i32;
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
                let occupied = self.reservations(&self.day).iter().any(|r| {
                    r.resource == RESOURCES[self.resource_idx].0
                        && r.start == hour
                        && r.status != "cancelada"
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
        let price = self.price_for_kind(RESOURCES[self.resource_idx].1);
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
                                "{}  ·  {:02}:00–{:02}:00",
                                RESOURCES[self.resource_idx].0,
                                self.start_hour,
                                self.start_hour + 1
                            ))
                            .strong()
                            .size(16.0),
                        );
                        ui.label(
                            RichText::new("1 hora · precio según tarifa vigente")
                                .small()
                                .color(Color32::GRAY),
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
                !self.customer.trim().is_empty(),
                egui::Button::new(RichText::new("CONFIRMAR RESERVA   →").strong())
                    .min_size(Vec2::new(260.0, 46.0))
                    .fill(Color32::from_rgb(29, 120, 137)),
            )
            .clicked()
        {
            self.create_reservation();
        }
        ui.separator();
        ui.heading(format!("Turnos de hoy · {}", self.day));
        self.reservation_list(ui);
    }
    fn reservation_list(&mut self, ui: &mut egui::Ui) {
        let rows = self.reservations(&self.day);
        egui::ScrollArea::vertical()
            .max_height(300.0)
            .show(ui, |ui| {
                for r in rows {
                    ui.horizontal(|ui| {
                        ui.label(format!("{:02}:00–{:02}:00", r.start, r.start + r.duration));
                        ui.label(&r.resource);
                        ui.strong(&r.customer);
                        ui.label(money(r.total));
                        ui.label(&r.status);
                    });
                }
            });
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
                                "◈\n{}\n{}\nStock {}",
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
        ui.collapsing("Cobrar una sesión reservada", |ui| {
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
                    if ui.button("Cobrar y cerrar sesión").clicked() {
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
        ui.label("Las nuevas reservas toman la tarifa vigente y conservan ese importe. Las ofertas y descuentos configurables vuelven en una próxima etapa.");
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
