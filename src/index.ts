import * as XLSX from 'xlsx';

/**
 * Excel to JSON with xlsx
 * @param filePath - path of excel file
 * @returns {any[]} - array of json object
 * function to convert excel file from path to json
 */
export const excelToJson = (filePath: string):object[] => {
    const workbook:XLSX.WorkBook = XLSX.readFile(filePath);
    const sheet_name_list:string[] = workbook.SheetNames;
    const xlData:object[] = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
    return xlData;
};

/**
 * Merge JSON with Same ID
 * @param {any[]} data - array of json object
 * @returns {any[]} - array of merged json object
 * Asumsi jika 1 excel terdiri dari beberapa objek
 * Merge the same keys with the same ID to array
 * Example:
 * [{"id":1, "name":"Josh"},{"id":1,"hobby":"swimming"},{"id":1,"hobby":"soccer"},{"id":2,"name":"john"}]
 * [{"id":1, "name":"Josh","hobby":["swimming","soccer"]},{"id":2,"name":"john"}]
 */
export const mergeWithSameID = (data: any[]): any[] => {
    const mergedData: any[] = [];

    for (const item of data) {
        const existingItem = mergedData.find((m) => m.id === item.id);

        if (existingItem) {
            // if the existing item has a property with the same name, append the value to the existing property
            Object.keys(item).forEach((key) => {
                if (key === "id") {
                    return;
                }
                if (existingItem[key]) {
                    if (Array.isArray(existingItem[key])) {
                        existingItem[key].push(item[key]);
                    } else {
                        existingItem[key] = [existingItem[key], item[key]];
                    }
                } else {
                    existingItem[key] = item[key];
                }
            });
        } else {
            mergedData.push(item);
        }
    }

    return mergedData;
}

/**
 * Merge Object with same keys to array
 * @param json - json from excel
 * @returns
 * Example:
 * [{"id":1, "name":"Josh"},{"id":1,"hobby":"swimming"},{"id":1,"hobby":"soccer"},{"id":1,"name":"john"}]
 * [{"id":1, "name":["Josh","john"],"hobby":["swimming","soccer"]}]
 */
export const mergeObjects = (json: any[]): any[] => {
    let result:any = {};
    json.forEach(item => {
        Object.keys(item).forEach(key => {
            if (!result[key]) {
                result[key] = [];
            }
            result[key].push(item[key]);
        });
    });
    Object.keys(result).forEach(key => {
        if (Array.isArray(result[key]) && result[key].length === 1) {
            result[key] = result[key][0];
        }
    });
    return [result];
}

/**
 * Unflatten Dot Notation JSON
 * @param data - json with dot notation key
 * @returns 
 */
export const unflattenDotJson = (data: any): any => {
    const result: any = {};
    for (const i in data) {
        const keys = i.split('.');
        keys.reduce((r, e, j) => {
            return r[e] || (r[e] = isNaN(Number(keys[j + 1])) ? (keys.length - 1 === j ? data[i] : {}) : []);
        }, result);
    }
    return result;
}

// call functions
const xlData = excelToJson('sample.xlsx');
console.log(xlData);
const mergedData = mergeObjects(xlData);
console.log(mergedData);
const unflattenData = unflattenDotJson(mergedData[0]);
console.log(unflattenData);

console.log("OK");
