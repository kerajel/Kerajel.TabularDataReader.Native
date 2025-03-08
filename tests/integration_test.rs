extern crate kerajel_tabular_data_reader;

use kerajel_tabular_data_reader::*;

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs::File;
    use std::io::Read;
    use std::ffi::{c_int, CStr, CString};

    #[test]
    fn test_excel_to_csv() {
        
        let mut file = File::open("tests/resources/test_data.xlsx").expect("file should open read only");
        let mut buffer = Vec::new();
        file.read_to_end(&mut buffer).expect("buffer should be filled");

        let sheet_name = CString::new("Sheet1").expect("CString::new failed");
        let sheet_name_ptr = sheet_name.as_ptr();

        let result = excel_to_csv(buffer.as_ptr(), buffer.len(), sheet_name_ptr);

        assert_eq!(result.operation_status, OperationStatus::Succeeded as c_int);

        assert!(!result.result.is_null());

        let csv_data = unsafe { 
            CStr::from_ptr(result.result)
                .to_str()
                .expect("Failed to convert to str")
                .to_string() 
        };

        let expected_csv = "\
\"id\",\"name\",\"color\"\n\
\"1\",\"orange\",\"orange\"\n\
\"2\",\"apple\",\"yellow\"\n\
\"3\",\"cucumber\",\"green\"\n\
\"4\",\"melon\",\"red\"\n\
\"5\",\"grape\",\"violet\"\n";
        

        assert_eq!(csv_data, expected_csv, "The CSV output does not match the expected data");

        println!("{}", csv_data);

        free_operation_result(result);
    }
}
