const fs = require('fs');
const path = require("path");
const XLSX = require("xlsx");
const fspromises = fs.promises;

const { getFileNameWithoutExtn } = require(`${ASBCONSTANTS.ROOTDIR}/custom/save_excel.js`);

exports.start = async (_routeName, _route, _messageContainer, message) => {
    if (message.content?.result) {
        const excelPath = message.content.excel_path;
        const excelToJsonResult = await _excel_to_json(excelPath);
        message.content = {...excelToJsonResult};
    }

    message.addRouteDone(_routeName);
    message.setGCEligible(true);
}

async function _excel_to_json(excel_path, save_csv=false) {
    const dir_to_save = path.dirname(excel_path);
    const excel_name = getFileNameWithoutExtn(excel_path);
    const destination_json = `${path.join(dir_to_save, excel_name)}.json`;
    try {
        const workbook = XLSX.readFile(excel_path);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        if(save_csv){
            const csv = XLSX.utils.sheet_to_csv(sheet);
            await fspromises.writeFile(`${path.join(dir_to_save, excel_name)}.csv`, csv);
        } const json = XLSX.utils.sheet_to_json(sheet);
        await fspromises.writeFile(destination_json, JSON.stringify(json, null, 4));
        const success_message = `Excel - ${excel_name} convert to JSON successfully!!`;
        return {result: true, message: success_message, json_path: destination_json};
    } catch (error) {
        const error_message = `Failed to convert Excel to JSON for the incoming Excel - ${excel_name}.`;
        ASBLOG.error(`Error in [${getFileNameWithoutExtn(__filename)}]: ${error_message} error: ${error.message}`);
        return {result: false, error: error_message};
    }
}