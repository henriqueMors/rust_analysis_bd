use calamine::{open_workbook, Reader, Xlsx};
use csv::ReaderBuilder;
use std::collections::HashSet;
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    // Lendo o arquivo CSV
    let mut firewalls_csv = ReaderBuilder::new()
        .delimiter(b',')
        .has_headers(true)
        .from_path("src/firewalls21jan25.csv")?;

    // Criar um HashSet com os nomes dos firewalls do CSV
    let mut firewalls_set = HashSet::new();
    for record in firewalls_csv.records() {
        let record = record?;
        if let Some(name) = record.get(1) {
            firewalls_set.insert(name.to_string());
        }
    }

    // Lista de arquivos Excel
    let excel_files = ["src/FW - MCK.xlsx", "src/FW - LPA.xlsx", "src/FW - CAS.xlsx"];

    // Iterar sobre os arquivos Excel e procurar os nomes
    for file in excel_files {
        let mut workbook: Xlsx<_> = open_workbook(file)?;

        for sheet in workbook.sheet_names().to_owned() {
            if let Some(Ok(range)) = workbook.worksheet_range(&sheet) {
                for row in range.rows() {
                    for cell in row {
                        if let Some(cell_value) = cell.get_string() {
                            if firewalls_set.contains(cell_value) {
                                println!(
                                    "Firewall encontrado: '{}' no arquivo '{}', planilha '{}'",
                                    cell_value, file, sheet
                                );
                            }
                        }
                    }
                }
            }
        }
    }

    Ok(())
}
