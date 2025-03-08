use calamine::{open_workbook_auto, open_workbook_auto_from_rs, Reader, Sheets};
use csv::{WriterBuilder, QuoteStyle};
use std::ffi::{CStr, CString};
use std::io::{Cursor, Read, Seek};
use std::os::raw::{c_char, c_int};
use std::ptr;

#[repr(C)]
#[derive(Debug)]
pub enum OperationStatus {
    Unknown = 1,
    Succeeded = 2,
    Faulted = 3,
}

#[repr(C)]
pub struct OperationResult {
    pub result: *mut c_char,
    pub operation_status: c_int,
    pub error_message: *mut c_char,
}

impl OperationResult {
    fn faulted(message: &str) -> OperationResult {
        let err_cstring = CString::new(message).unwrap_or_else(|_| {
            CString::new("Failed to allocate error message")
                .expect("Fallback CString conversion in faulted()")
        });

        OperationResult {
            result: ptr::null_mut(),
            operation_status: OperationStatus::Faulted as c_int,
            error_message: err_cstring.into_raw(),
        }
    }

    fn succeeded(result_cstring: CString) -> OperationResult {
        OperationResult {
            result: result_cstring.into_raw(),
            operation_status: OperationStatus::Succeeded as c_int,
            error_message: ptr::null_mut(),
        }
    }
}

fn convert_sheet_to_csv<R: Read + Seek>(
    mut workbook: Sheets<R>,
    sheet_name: *const c_char
) -> OperationResult {
    let selected_sheet = if sheet_name.is_null() {
        workbook.sheet_names().get(0).cloned()
    } else {
        let sheet_name_result = unsafe { CStr::from_ptr(sheet_name).to_str() };
        match sheet_name_result {
            Ok(s) => Some(s.to_string()),
            Err(_) => return OperationResult::faulted("Invalid UTF-8 in sheet name"),
        }
    };

    let range = match selected_sheet {
        Some(name) => match workbook.worksheet_range(&name) {
            Ok(rng) => rng,
            Err(_) => return OperationResult::faulted("Error accessing specified sheet"),
        },
        None => return OperationResult::faulted(
            "Sheet name is null and no sheets are available"
        ),
    };

    let mut wtr = WriterBuilder::new()
        .quote_style(QuoteStyle::Always)
        .from_writer(vec![]);

    for row in range.rows() {
        let row_data: Vec<String> = row.iter().map(|cell| cell.to_string()).collect();
        if let Err(_) = wtr.write_record(&row_data) {
            return OperationResult::faulted("Failed to write CSV record");
        }
    }

    let csv_bytes = match wtr.into_inner() {
        Ok(b) => b,
        Err(_) => return OperationResult::faulted("Failed to retrieve CSV data"),
    };

    let csv_string = match String::from_utf8(csv_bytes) {
        Ok(s) => s,
        Err(_) => return OperationResult::faulted("CSV data not valid UTF-8"),
    };

    match CString::new(csv_string) {
        Ok(cstr) => OperationResult::succeeded(cstr),
        Err(_) => OperationResult::faulted("Failed to create C string"),
    }
}

#[no_mangle]
pub extern "C" fn excel_to_csv(
    bytes: *const u8,
    len: usize,
    sheet_name: *const c_char,
) -> OperationResult {
    if bytes.is_null() {
        return OperationResult::faulted("Null pointer received for bytes");
    }

    let slice = unsafe { std::slice::from_raw_parts(bytes, len) };
    let cursor = Cursor::new(slice);

    let workbook = match open_workbook_auto_from_rs(cursor) {
        Ok(wb) => wb,
        Err(e) => {
            let err_str = format!("Failed to open workbook: {}", e);
            return OperationResult::faulted(&err_str);
        }
    };

    convert_sheet_to_csv(workbook, sheet_name)
}

#[no_mangle]
pub extern "C" fn excel_to_csv_by_path(
    path: *const c_char,
    sheet_name: *const c_char,
) -> OperationResult {
    if path.is_null() {
        return OperationResult::faulted("Null pointer for file path");
    }

    let path_str = unsafe {
        match CStr::from_ptr(path).to_str() {
            Ok(s) => s.to_string(),
            Err(_) => return OperationResult::faulted("Invalid UTF-8 in file path"),
        }
    };

    let workbook = match open_workbook_auto(&path_str) {
        Ok(wb) => wb,
        Err(e) => {
            let err_str = format!("Cannot open file: {}", e);
            return OperationResult::faulted(&err_str);
        }
    };

    convert_sheet_to_csv(workbook, sheet_name)
}

#[no_mangle]
pub extern "C" fn free_operation_result(result: OperationResult) {
    unsafe {
        if !result.result.is_null() {
            let _ = CString::from_raw(result.result);
        }
        if !result.error_message.is_null() {
            let _ = CString::from_raw(result.error_message);
        }
    }
}