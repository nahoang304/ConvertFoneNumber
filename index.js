// shiva8x@gmail.com
// 20190127 - 1.0.0 read txt, convert to xls
// 20190207 - 1.1 read xlsx, xls, convert to xlsx
const xlsx = require('xlsx');
const path = require('path');
let lines = [];

const params = require('minimist')(process.argv.slice(2));
console.log("params: " + JSON.stringify(params, undefined, 3));
let file_path = (params["_"] ? params["_"] : undefined);
console.log("file_path: " + file_path);
let output_file_path = "";
if (file_path === undefined || file_path.length === 0) {
    console.log("Error: Invalid File Path!")
    return
}
file_path = file_path + ""
let extension = path.extname(file_path);
let base_name = path.basename(file_path, extension);
// console.log("base_name: " + base_name);
output_file_path = "./output/" + base_name + "_proccessed" + extension;
// console.log("output_file_path: " + output_file_path);
let workbook = xlsx.readFile(file_path);
let worksheet = workbook.Sheets[workbook.SheetNames[0]];
// console.log("worksheet: " + JSON.stringify(worksheet, undefined, 3));
let range = xlsx.utils.decode_range(worksheet['!ref']);
// console.log("range: " + JSON.stringify(range, undefined, 3));
// load data
for (var i = range.s.r; i <= range.e.r; ++i) {
    for (var j = range.s.c; j <= range.e.c; ++j) {
        let cell_address = {c:j, r:i};
        let cell_ref = xlsx.utils.encode_cell(cell_address);
        // console.log("cell_ref: " + JSON.stringify(cell_ref, undefined, 3));
        let cell = worksheet[cell_ref];
        // console.log("cell: " + JSON.stringify(cell, undefined, 3));
        let cell_value = (cell ? cell.w : undefined);
        if (cell_value !== undefined && cell_value.length > 0) {
            // console.log("cell value: " + cell_value);
            // remove all whitespaces, special characters: " ", "."
            let line = cell_value.replace(/[^0-9]/g,"");
            // processing
            if (line.startsWith("0") && line.length === 11) { // 01689 113 113
                let fourDigits = line.substr(0,4) // 0168
                switch (fourDigits) {
                    // viettel
                    case "0162":
                        line = line.replace("0162", "32")
                        break;
                    case "0163":
                        line = line.replace("0163", "33")
                        break;
                    case "0164":
                        line = line.replace("0164", "34")
                        break;
                    case "0165":
                        line = line.replace("0165", "35")
                        break;
                    case "0166":
                        line = line.replace("0166", "36")
                        break;
                    case "0167":
                        line = line.replace("0167", "37")
                        break;
                    case "0168":
                        line = line.replace("0168", "38")
                        break;
                    case "0169":
                        line = line.replace("0169", "39")
                        break;
                    // vinaphone
                    case "0123":
                        line = line.replace("0123", "83")
                        break;
                    case "0124":
                        line = line.replace("0124", "84")
                        break;
                    case "0125":
                        line = line.replace("0125", "85")
                        break;
                    case "0127":
                        line = line.replace("0127", "81")
                        break;
                    case "0129":
                        line = line.replace("0129", "82")
                        break;
                    // mobiphone
                    case "0120":
                        line = line.replace("0120", "70")
                        break;
                    case "0121":
                        line = line.replace("0121", "79")
                        break;
                    case "0122":
                        line = line.replace("0122", "77")
                        break;
                    case "0126":
                        line = line.replace("0126", "76")
                        break;
                    case "0128":
                        line = line.replace("0128", "78")
                        break;
                    // vietnamobile
                    case "0186":
                        line = line.replace("0186", "56")
                        break;
                    case "0188":
                        line = line.replace("0188", "58")
                        break;
                    // G-mobile
                    case "0199":
                        line = line.replace("0199", "59")
                        break;
                    default:
                        break;
                }
            } else if(line.startsWith("84") && line.length === 12) { // 841689 113 113
                let fiveDigits = line.substr(0,5) // 84168
                switch (fiveDigits) {
                    // viettel
                    case "84162":
                        line = line.replace("84162", "32")
                        break;
                    case "84163":
                        line = line.replace("84163", "33")
                        break;
                    case "84164":
                        line = line.replace("84164", "34")
                        break;
                    case "84165":
                        line = line.replace("84165", "35")
                        break;
                    case "84166":
                        line = line.replace("84166", "36")
                        break;
                    case "84167":
                        line = line.replace("84167", "37")
                        break;
                    case "84168":
                        line = line.replace("84168", "38")
                        break;
                    case "84169":
                        line = line.replace("84169", "39")
                        break;
                    // vinaphone
                    case "84123":
                        line = line.replace("84123", "83")
                        break;
                    case "84124":
                        line = line.replace("84124", "84")
                        break;
                    case "84125":
                        line = line.replace("84125", "85")
                        break;
                    case "84127":
                        line = line.replace("84127", "81")
                        break;
                    case "84129":
                        line = line.replace("84129", "82")
                        break;
                    // mobiphone
                    case "84120":
                        line = line.replace("84120", "70")
                        break;
                    case "84121":
                        line = line.replace("84121", "79")
                        break;
                    case "84122":
                        line = line.replace("84122", "77")
                        break;
                    case "84126":
                        line = line.replace("84126", "76")
                        break;
                    case "84128":
                        line = line.replace("84128", "78")
                        break;
                    // vietnamobile
                    case "84186":
                        line = line.replace("84186", "56")
                        break;
                    case "84188":
                        line = line.replace("84188", "58")
                        break;
                    // G-mobile
                    case "84199":
                        line = line.replace("84199", "59")
                        break;
                    default:
                        break;
                }
            } else if (line.startsWith("84") && line.length === 11) { // 84988123456
                line = line.substr(2,line.length)
            } else if (line.startsWith("0") && line.length === 10) { // 0912345678
                line = line.substr(1,line.length)
            } else if (line.startsWith("1") &&  line.length === 10) { // 1685 229 464
                let threeDigits = line.substr(0,3) // 168
                switch (threeDigits) {
                    // viettel
                    case "162":
                        line = line.replace("162", "32")
                        break;
                    case "163":
                        line = line.replace("163", "33")
                        break;
                    case "164":
                        line = line.replace("164", "34")
                        break;
                    case "165":
                        line = line.replace("165", "35")
                        break;
                    case "166":
                        line = line.replace("166", "36")
                        break;
                    case "167":
                        line = line.replace("167", "37")
                        break;
                    case "168":
                        line = line.replace("168", "38")
                        break;
                    case "169":
                        line = line.replace("169", "39")
                        break;
                    // vinaphone
                    case "123":
                        line = line.replace("123", "83")
                        break;
                    case "124":
                        line = line.replace("124", "84")
                        break;
                    case "125":
                        line = line.replace("125", "85")
                        break;
                    case "127":
                        line = line.replace("127", "81")
                        break;
                    case "129":
                        line = line.replace("129", "82")
                        break;
                    // mobiphone
                    case "120":
                        line = line.replace("120", "70")
                        break;
                    case "121":
                        line = line.replace("121", "79")
                        break;
                    case "122":
                        line = line.replace("122", "77")
                        break;
                    case "126":
                        line = line.replace("126", "76")
                        break;
                    case "128":
                        line = line.replace("128", "78")
                        break;
                    // vietnamobile
                    case "186":
                        line = line.replace("186", "56")
                        break;
                    case "188":
                        line = line.replace("188", "58")
                        break;
                    // G-mobile
                    case "199":
                        line = line.replace("199", "59")
                        break;
                    default:
                        break;
                }
            } else {
                // should never happen, just leave invalid data alone
            }
            // finished
            if (line.length > 0) {
                lines.push(line)
                console.log("Pushed line: " + line)
            }
        }
    }
}
// write data
if (lines.length > 0) {
    let ws_name = "Finished";
    let ws_data = [];
    let column = Array.from(lines);
    let aoa = [];
    for(var i = 0; i < 1; ++i) {
      for(var j = 0; j < lines.length; ++j) {
        if(!aoa[j]) aoa[j] = [];
        aoa[j][i] = column[j];
      }
    }
    let ws = xlsx.utils.aoa_to_sheet(aoa);
    let wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, ws_name);
    xlsx.writeFile(wb, output_file_path);
    console.log("== Finished " + lines.length + " records ==")
}
