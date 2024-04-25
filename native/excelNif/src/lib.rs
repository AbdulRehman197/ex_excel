use calamine::{open_workbook, Reader, Xlsx,Ods};

#[rustler::nif]
fn xlsx(file_path: &str,sheet_name: &str) -> Result<Vec<Vec<String>>, String> {
    // Try to open the workbook
    let mut workbook: Xlsx<_> = match open_workbook(file_path) {
        Ok(workbook) => workbook,
        Err(e) => return Err(format!("Failed to open workbook: {}", e)),
    };

    // Get the first sheet

    let sheet = match workbook.worksheet_range(sheet_name){
        Ok(sheet) => sheet,
        Err(e) => return Err(format!("Failed to read sheet: {}", e)),
      
    };

    // Convert the sheet data into a vector of vectors of strings
    let mut excel_data = Vec::new();
    for row in sheet.rows() {
        let mut excel_row = Vec::new();
        for cell in row {
                excel_row.push(cell.to_string());
        }
        excel_data.push(excel_row);
    }

    Ok(excel_data)
}

#[rustler::nif]
fn ods(file_path: &str,sheet_name: &str) -> Result<Vec<Vec<String>>, String> {
    // Try to open the workbook
    let mut workbook: Ods<_> = match open_workbook(file_path) {
        Ok(workbook) => workbook,
        Err(e) => return Err(format!("Failed to open workbook: {}", e)),
    };

    // Get the first sheet

    let sheet = match workbook.worksheet_range(sheet_name){
        Ok(sheet) => sheet,
        Err(e) => return Err(format!("Failed to read sheet: {}", e)),
      
    };

    // Convert the sheet data into a vector of vectors of strings
    let mut excel_data = Vec::new();
    for row in sheet.rows() {
        let mut excel_row = Vec::new();
        for cell in row {
                excel_row.push(cell.to_string());
        }
        excel_data.push(excel_row);
    }

    Ok(excel_data)
}
#[rustler::nif]
fn xlsx_sheet_names (file_path: &str) -> Result<Vec<String>, String> {
    let  workbook: Xlsx<_> = match open_workbook(file_path) {
        Ok(workbook) => workbook,
        Err(e) => return Err(format!("Failed to open workbook: {}", e)),
    };
    return Ok(workbook.sheet_names());
}

#[rustler::nif]
fn ods_sheet_names (file_path: &str) -> Result<Vec<String>, String> {
    let  workbook: Ods<_> = match open_workbook(file_path) {
        Ok(workbook) => workbook,
        Err(e) => return Err(format!("Failed to open workbook: {}", e)),
    };
    return Ok(workbook.sheet_names());
}
rustler::init!("Elixir.ExExcel", [xlsx,ods, xlsx_sheet_names,ods_sheet_names]);
