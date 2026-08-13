const fs = require('fs');
const path = require("path");
const fspromises = fs.promises;

const utils = require(`${ASBCONSTANTS.LIBDIR}/utils.js`);

exports.start = async (_routeName, _route, _messageContainer, message) => {
    const excelData = message.content.data, excelName = message.content.filename;
    const excelExtn = message.content.extension || "xlsx";
    const targetDir = _route.dir_to_save || path.join(ASBCONSTANTS.ROOTDIR, "in_excels");

    const saveExcelResult = await _saveExcel(targetDir, excelName, excelData, excelExtn, "base64");

    message.content = {...saveExcelResult};
    message.addRouteDone(_routeName);
    message.setGCEligible(true);
}

async function _saveExcel(dir_to_save, excel_name, excel_content, excel_extn, content_type) {
    if (!excel_name || !excel_content) return {return: false, 
        error: "Invalid arguments!!\nRequired: filename & data\nOptional: extension"};
    const destination_path = `${path.join(dir_to_save, excel_name)}_${utils.getTimeStamp()}.${excel_extn}`;
    try {
        await fspromises.mkdir(dir_to_save, {recursive: true});
        await fspromises.writeFile(destination_path, excel_content, content_type);
        const success_message = `Excel - ${excel_name} saved successfully!!`;
        return {result: true, message: success_message, excel_path: destination_path};
    } catch (error) {
        const error_message = `Failed to save the incoming Excel - ${excel_name}.`;
        ASBLOG.error(`Error in [${exports.getFileNameWithoutExtn(__filename)}]: ${error_message} error: ${error.message}`);
        return {result: false, error: error_message};
    }
}

exports.getFileNameWithoutExtn = file => path.basename(file).split(".").slice(0, -1).join(".");