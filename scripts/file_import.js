// Text/DataTable/DataReader
// function sample_method(value, record)
// {
//     return value + record.GetValue(column_name);
// }
// Excel
// function sample_method(excel, value, record)
// {
//     1) return value + record.GetValue(column_name);
//     2) return excel.Range("Sheet1", "A1");
// }
//
// ADD
// function sample_method(value)
// {
//        return value + "add";
// }

function gender(value, record) {
    if (value == "1") {
        return "男";
    }
    else if (value == "2") {
        return "女";
    }
    else {
        return "";
    }
}

function add_col(value, recoed)
{
    return value + "：JavaScript";
}
