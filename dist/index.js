"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.reformatChildObject = exports.unflattenDotJson = exports.mergeObjects = exports.mergeWithSameID = exports.excelToJson = void 0;
const XLSX = __importStar(require("xlsx"));
const excelToJson = (filePath) => {
    const workbook = XLSX.readFile(filePath);
    const sheet_name_list = workbook.SheetNames;
    const xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
    return xlData;
};
exports.excelToJson = excelToJson;
const mergeWithSameID = (data) => {
    const mergedData = [];
    for (const item of data) {
        const existingItem = mergedData.find((m) => m.id === item.id);
        if (existingItem) {
            Object.keys(item).forEach((key) => {
                if (key === "id") {
                    return;
                }
                if (existingItem[key]) {
                    if (Array.isArray(existingItem[key])) {
                        existingItem[key].push(item[key]);
                    }
                    else {
                        existingItem[key] = [existingItem[key], item[key]];
                    }
                }
                else {
                    existingItem[key] = item[key];
                }
            });
        }
        else {
            mergedData.push(item);
        }
    }
    return mergedData;
};
exports.mergeWithSameID = mergeWithSameID;
const mergeObjects = (json) => {
    let result = {};
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
    return result;
};
exports.mergeObjects = mergeObjects;
const unflattenDotJson = (data) => {
    const result = {};
    for (const i in data) {
        const keys = i.split('.');
        keys.reduce((r, e, j) => {
            return r[e] || (r[e] = isNaN(Number(keys[j + 1])) ? (keys.length - 1 === j ? data[i] : {}) : []);
        }, result);
    }
    return result;
};
exports.unflattenDotJson = unflattenDotJson;
const reformatChildObject = (obj) => {
    const result = {};
    for (const key in obj) {
        if (typeof obj[key] === 'object') {
            const keys = Object.keys(obj[key]);
            const values = Object.values(obj[key]);
            const length = values[0].length;
            if (!values.every(Array.isArray) || !values.every(v => v.length === length)) {
                throw new Error('Invalid data');
            }
            const temp = [];
            for (let j = 0; j < length; j++) {
                const obj = {};
                for (let k = 0; k < keys.length; k++) {
                    obj[keys[k]] = values[k][j];
                }
                temp.push(obj);
            }
            result[key] = temp;
        }
        else {
            result[key] = obj[key];
        }
    }
    return result;
};
exports.reformatChildObject = reformatChildObject;
const xlData = (0, exports.excelToJson)('client-side/sample.xlsx');
console.log(xlData);
const mergedData = (0, exports.mergeObjects)(xlData);
console.log(mergedData);
const unflattenData = (0, exports.unflattenDotJson)(mergedData);
console.log(unflattenData);
const reformatData = (0, exports.reformatChildObject)(unflattenData);
console.log(reformatData);
console.log("OK");
//# sourceMappingURL=index.js.map