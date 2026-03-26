use calamine::{open_workbook_auto, Data, Reader};
use std::fs::File;
use rust_xlsxwriter::{Color, Format, Workbook};
use std::collections::{HashMap, HashSet};
use std::io::{self, Write};
use std::path::Path;

// ─── helpers ─────────────────────────────────────────────────────────────────

fn prompt(msg: &str) -> String {
    print!("{msg}");
    io::stdout().flush().unwrap();
    let mut buf = String::new();
    io::stdin().read_line(&mut buf).unwrap();
    buf.trim().trim_matches('"').trim_matches('\'').to_string()
}

fn cell_to_string(cell: &Data) -> Option<String> {
    match cell {
        Data::Empty => None,
        Data::String(s) => {
            let n = s.trim().to_lowercase();
            if n.is_empty() { None } else { Some(n) }
        }
        Data::Float(f) => Some(format!("{f}")),
        Data::Int(i)   => Some(format!("{i}")),
        Data::Bool(b)  => Some(format!("{b}")),
        Data::DateTime(dt) => Some(format!("{dt}")),
        _ => None,
    }
}

// ─── load workbook ───────────────────────────────────────────────────────────

fn load_csv(path: &str) -> Option<(Vec<Vec<Option<String>>>, Vec<String>)> {
    let file = match File::open(path) {
        Ok(f) => f,
        Err(e) => { println!("  Could not open file: {e}\n"); return None; }
    };
    let mut rdr = csv::ReaderBuilder::new().delimiter(b';').from_reader(file);
    let headers: Vec<String> = match rdr.headers() {
        Ok(h) => h.iter().map(|s| s.trim().trim_start_matches('\u{FEFF}').to_lowercase()).collect(),
        Err(e) => { println!("  Could not read CSV headers: {e}\n"); return None; }
    };
    let data: Vec<Vec<Option<String>>> = rdr.records()
        .filter_map(|r| r.ok())
        .map(|row| row.iter().map(|v| {
            let s = v.trim().to_lowercase();
            if s.is_empty() { None } else { Some(s) }
        }).collect())
        .collect();
    println!("  Loaded {} rows. Columns: {headers:?}\n", data.len());
    Some((data, headers))
}

fn load_file(label: &str, path: &str) -> (Vec<Vec<Option<String>>>, Vec<String>) {
    println!("Loading {label}: {path}");
    if !Path::new(path).exists() {
        panic!("  File not found: {path}");
    }

    let ext = Path::new(path).extension()
        .and_then(|e| e.to_str())
        .unwrap_or("")
        .to_lowercase();

    if ext == "csv" {
        return load_csv(path).expect("  Failed to load CSV");
    }

    let mut wb = open_workbook_auto(path)
        .unwrap_or_else(|e| panic!("  Could not open file: {e}"));

    let sheet_names = wb.sheet_names().to_vec();
    println!("  Sheets: {sheet_names:?}");
    let sheet_input = prompt("  Sheet name (blank = first sheet): ");
    let sheet_name = if sheet_input.is_empty() {
        sheet_names[0].clone()
    } else {
        sheet_input
    };

    let range = wb.worksheet_range(&sheet_name)
        .unwrap_or_else(|e| panic!("  Could not read sheet: {e}"));

    let mut rows_iter = range.rows();

    let headers: Vec<String> = match rows_iter.next() {
        Some(row) => row.iter()
            .map(|c| cell_to_string(c).unwrap_or_default())
            .collect(),
        None => panic!("  Sheet is empty."),
    };

    let data: Vec<Vec<Option<String>>> = rows_iter
        .map(|row| row.iter().map(cell_to_string).collect())
        .collect();

    println!("  Loaded {} rows. Columns: {headers:?}\n", data.len());
    (data, headers)
}

// ─── pick column ─────────────────────────────────────────────────────────────

fn pick_column(headers: &[String], label: &str) -> usize {
    println!("Columns in {label} file:");
    for (i, h) in headers.iter().enumerate() {
        println!("  [{i}] {h}");
    }
    loop {
        let choice = prompt("Select column (name or index): ");
        if let Ok(idx) = choice.parse::<usize>() {
            if idx < headers.len() { return idx; }
            println!("  Index out of range.");
        } else {
            let lower = choice.to_lowercase();
            if let Some(idx) = headers.iter().position(|h| h.to_lowercase() == lower) {
                return idx;
            }
            println!("  Column not found, try again.");
        }
    }
}

// ─── extract column values ────────────────────────────────────────────────────

fn extract_column(data: &[Vec<Option<String>>], col_idx: usize) -> Vec<Option<String>> {
    data.iter()
        .map(|row| row.get(col_idx).cloned().flatten())
        .collect()
}

// ─── partial match ───────────────────────────────────────────────────────────

fn keywords(s: &str) -> HashSet<String> {
    s.split(|c: char| !c.is_alphabetic())
        .filter(|w| w.len() > 2)
        .map(|w| w.to_lowercase())
        .collect()
}

fn partial_match(a: &str, b: &str) -> bool {
    if a.contains(b) || b.contains(a) { return true; }
    keywords(a).intersection(&keywords(b)).next().is_some()
}

// ─── compare & print stats ────────────────────────────────────────────────────

struct Stats {
    col_a_name:    String,
    col_b_name:    String,
    common:        Vec<String>,
    only_in_a:     Vec<String>,
    only_in_b:     Vec<String>,
    values_a:      Vec<Option<String>>,
    set_b:         HashSet<String>,
    matched_pairs: Vec<(String, String)>,
}

fn compare(
    col_a_name: &str,
    values_a: &[Option<String>],
    col_b_name: &str,
    values_b: &[Option<String>],
) -> Stats {
    let valid_a: Vec<&str> = values_a.iter().flatten().map(|s| s.as_str()).collect();
    let valid_b: Vec<&str> = values_b.iter().flatten().map(|s| s.as_str()).collect();

    let set_a: HashSet<&str> = valid_a.iter().copied().collect();
    let set_b: HashSet<&str> = valid_b.iter().copied().collect();

    // build matched pairs: (value_a, value_b)
    let mut matched_pairs: Vec<(String, String)> = set_a.iter()
        .flat_map(|a| {
            set_b.iter()
                .filter(|b| partial_match(a, b))
                .map(|b| (a.to_string(), b.to_string()))
                .collect::<Vec<_>>()
        })
        .collect();
    matched_pairs.sort();

    let mut common:    Vec<String> = set_a.iter()
        .filter(|a| set_b.iter().any(|b| partial_match(a, b)))
        .map(|s| s.to_string()).collect();
    let mut only_in_a: Vec<String> = set_a.iter()
        .filter(|a| !set_b.iter().any(|b| partial_match(a, b)))
        .map(|s| s.to_string()).collect();
    let mut only_in_b: Vec<String> = set_b.iter()
        .filter(|b| !set_a.iter().any(|a| partial_match(a, b)))
        .map(|s| s.to_string()).collect();
    common.sort();
    only_in_a.sort();
    only_in_b.sort();

    let rows_a_total   = valid_a.len();
    let rows_a_matched = valid_a.iter().filter(|a| set_b.iter().any(|b| partial_match(a, b))).count();
    let hit_rate  = if set_a.is_empty() { 0.0 } else { common.len() as f64 / set_a.len() as f64 * 100.0 };
    let row_rate  = if rows_a_total == 0 { 0.0 } else { rows_a_matched as f64 / rows_a_total as f64 * 100.0 };

    println!("\n{}", "=".repeat(55));
    println!("  COMPARISON RESULTS");
    println!("{}", "=".repeat(55));
    println!("  Basis column : '{col_a_name}'");
    println!("  Ziel column  : '{col_b_name}'");
    println!("{}", "-".repeat(55));
    println!("  Unique values in Basis        : {}", set_a.len());
    println!("  Unique values in Ziel         : {}", set_b.len());
    println!("  Common (Basis ∩ Ziel)         : {}", common.len());
    println!("  Only in Basis                 : {}", only_in_a.len());
    println!("  Only in Ziel                  : {}", only_in_b.len());
    println!("{}", "-".repeat(55));
    println!("  Hit rate  (Basis in Ziel)     : {hit_rate:.1}%");
    println!("  Row-level match rate (Basis)  : {row_rate:.1}%  ({rows_a_matched} / {rows_a_total} rows)");
    println!("{}", "=".repeat(55));

    println!("\n  MATCHED PAIRS (Basis → Ziel)");
    println!("{}", "-".repeat(55));
    for (a, b) in &matched_pairs {
        println!("  {a}  →  {b}");
    }
    println!("{}", "=".repeat(55));

    println!("\n  NO MATCH IN ZIEL (only in Basis)");
    println!("{}", "-".repeat(55));
    for v in &only_in_a {
        println!("  {v}");
    }
    println!("{}", "=".repeat(55));

    println!("\n  NO MATCH IN BASIS (only in Ziel)");
    println!("{}", "-".repeat(55));
    for v in &only_in_b {
        println!("  {v}");
    }
    println!("{}", "=".repeat(55));

    Stats {
        col_a_name: col_a_name.to_string(),
        col_b_name: col_b_name.to_string(),
        common,
        only_in_a,
        only_in_b,
        values_a: values_a.to_vec(),
        set_b: set_b.into_iter().map(|s| s.to_string()).collect(),
        matched_pairs,
    }
}

// ─── export ───────────────────────────────────────────────────────────────────

fn export_all(
    all_stats: &[(Stats, usize)],
    headers_a: &[String],
    data_a: &[Vec<Option<String>>],
    filename_idx: Option<usize>,
) {
    let drop_cols: HashSet<&str> = [
        "parent8","parent9","parent10","parent11","parent12","parent13",
    ].iter().copied().collect();

    let rename_col = |h: &str| -> String {
        match h {
            "parent4" => "Gruppe".to_string(),
            "parent5" => "Vorgang".to_string(),
            "parent6" => "Dokumentenart".to_string(),
            "parent7" => "Kostengruppe".to_string(),
            other => other.to_string(),
        }
    };

    let out_cols: Vec<usize> = headers_a.iter().enumerate()
        .filter(|(_, h)| !drop_cols.contains(h.as_str()))
        .map(|(i, _)| i)
        .collect();

    let out_headers: Vec<String> = out_cols.iter()
        .map(|&i| rename_col(&headers_a[i]))
        .collect();

    let stats_map: HashMap<usize, &Stats> = all_stats.iter()
        .map(|(s, idx)| (*idx, s))
        .collect();

    let mut wb = Workbook::new();
    let fmt_green = Format::new().set_background_color(Color::RGB(0x92D050));
    let fmt_red   = Format::new().set_background_color(Color::RGB(0xFF6666));

    let ws = wb.add_worksheet();
    ws.set_name("Ergebnis").unwrap();

    for (c, h) in out_headers.iter().enumerate() {
        ws.write(0, c as u16, h.as_str()).unwrap();
    }

    for (r, row) in data_a.iter().enumerate() {
        let row_num = (r + 1) as u32;
        for (c, &col_idx) in out_cols.iter().enumerate() {
            let val = row.get(col_idx).and_then(|v| v.as_deref()).unwrap_or("");
            let col_num = c as u16;

            if let Some(stats) = stats_map.get(&col_idx) {
                if val.is_empty() {
                    ws.write(row_num, col_num, "").unwrap();
                } else {
                    let is_filename = filename_idx.map(|fi| {
                        let raw = row.get(fi).and_then(|v| v.as_deref()).unwrap_or("");
                        let stripped = match raw.rfind('.') {
                            Some(pos) => &raw[..pos],
                            None => raw,
                        };
                        val == stripped
                    }).unwrap_or(false);

                    let matched = if is_filename {
                        None
                    } else {
                        stats.set_b.iter().find(|b| partial_match(val, b))
                    };
                    match matched {
                        Some(ziel_val) => {
                            ws.write_with_format(row_num, col_num, ziel_val.as_str(), &fmt_green).unwrap();
                        }
                        None => {
                            ws.write_with_format(row_num, col_num, val, &fmt_red).unwrap();
                        }
                    }
                }
            } else {
                ws.write(row_num, col_num, val).unwrap();
            }
        }
    }

    let out = "comparison_results.xlsx";
    wb.save(out).unwrap();
    println!("  Saved → {out}");
}

// ─── main ─────────────────────────────────────────────────────────────────────

const BASE_STRUCT: &str = r"C:\Users\ADM1.E.Baack\Desktop\pe_migration\pe_1274_docs.csv";
const GOAL_SRUCT: &str = r"C:\Users\ADM1.E.Baack\Desktop\pe_migration\2026-03-05_PE_Ablage_aufbereitet.csv";

fn main() {
    println!("=== Excel Column Comparison Tool ===\n");

    let (data_a, headers_a) = load_file("Basis", BASE_STRUCT);
    let (data_b, headers_b) = load_file("Ziel", GOAL_SRUCT);

    let filename_idx_a = headers_a.iter().position(|h| h == "filename");

    let mut column_pairs: Vec<(usize, usize)> = Vec::new();
    loop {
        println!("\n--- Pick column from Basis ---");
        let col_a_idx = pick_column(&headers_a, "Basis");

        println!("\n--- Pick column from Ziel ---");
        let col_b_idx = pick_column(&headers_b, "Ziel");

        column_pairs.push((col_a_idx, col_b_idx));

        let more = prompt("\nAdd another column pair? (y/n): ");
        if !more.eq_ignore_ascii_case("y") { break; }
    }

    let mut all_stats: Vec<(Stats, usize)> = Vec::new();

    for (col_a_idx, col_b_idx) in &column_pairs {
        let col_a_idx = *col_a_idx;
        let col_b_idx = *col_b_idx;

        let values_a: Vec<Option<String>> = data_a.iter().map(|row| {
            let val = row.get(col_a_idx).cloned().flatten();
            if let (Some(v), Some(fi)) = (&val, filename_idx_a) {
                let filename_raw = row.get(fi).and_then(|f| f.as_deref()).unwrap_or("");
                let filename_val = match filename_raw.rfind('.') {
                    Some(pos) => &filename_raw[..pos],
                    None => filename_raw,
                };
                if v.as_str() == filename_val {
                    return None;
                }
            }
            val
        }).collect();

        let values_b = extract_column(&data_b, col_b_idx);

        let stats = compare(
            &headers_a[col_a_idx],
            &values_a,
            &headers_b[col_b_idx],
            &values_b,
        );

        all_stats.push((stats, col_a_idx));
    }

    let choice = prompt("\nExport all results to Excel? (y/n): ");
    if choice.eq_ignore_ascii_case("y") {
        export_all(&all_stats, &headers_a, &data_a, filename_idx_a);
    }
}
