const fs = require('fs');
const path = require("path");
const XLSX = require("xlsx");
const fspromises = fs.promises;
const crypto = require("crypto");
const sqlite3 = require("sqlite3");
const mustache = require('mustache');
const mkdirAsync = _promisify(fs.mkdir);
const accessAsync = _promisify(fs.access);

const crypt = require(`${ASBCONSTANTS.LIBDIR}/crypt.js`);
const { getFileNameWithoutExtn } = require(`${ASBCONSTANTS.ROOTDIR}/custom/save_excel.js`);

let db, dbRunAsync;

exports.start = async (_routeName, _route, _messageContainer, message) => {
    if (message.content?.result) {
        const jsonPath = message.content.json_path;
        const dbPath = path.resolve(_route.db_path);
        const dbCreationSQLsPath = path.resolve(_route.db_creation_sqls_path);
        const dbCreationSQLs = require(dbCreationSQLsPath);
        const templatePath = path.resolve(_route.template_path);
        const template = require(templatePath);
        const saveMonsResult = await _save_mons_to_db(jsonPath, dbPath, dbCreationSQLs, template);
        message.content = {...saveMonsResult};
    }

    message.addRouteDone(_routeName);
    message.setGCEligible(true);
}

async function _save_mons_to_db(json_path, db_path, db_creation_sqls, template) {
    const rows = require(json_path);
    const dir_to_save = path.dirname(json_path);
    const json_name = getFileNameWithoutExtn(json_path);
    const destination_mons = `${path.join(dir_to_save, json_name)}_mon.json`;
    if(!await _initDB(db_path, db_creation_sqls)) return {
        result: false, message: "Failed to extract & store the mons due to internal issue." }
    try {
        let mons = {}; for (const row of rows) {
            const nodeInfo = _getNodeInfo(row);
            await _storeNodeToDB(db, nodeInfo);
            const monInfos = await _storeAndgetMonInfos(db, nodeInfo.node_id, row);
            const monJsons = _getMonsJson(nodeInfo, monInfos, template);
            mons = {...mons, ...monJsons};
        } await fspromises.writeFile(destination_mons, JSON.stringify(mons, null, 4));
        const success_message = `All mons are loaded to the db from Excel successfully!!`;
        return {result: true, message: success_message, mons_path: destination_mons};
    } catch (error) {
        const error_message = `Failed to extract & store the mons from Excel.`;
        ASBLOG.error(`Error in [${getFileNameWithoutExtn(__filename)}]: ${error_message} error: ${error.message}`);
        return {result: false, error: error_message};
    }
}

function _getMonsJson(node_info, mon_infos, template) {
    const mons = {}; for (const mon_info of mon_infos) {
        let mon_template = template[mon_info.type][mon_info.name];
        const state_dir_path = `${node_info.user!=="root"?"/home":""}/${node_info.user}/monboss_stats`;
        const decrypted_password = crypt.decrypt(node_info.password);
        const data = {...node_info, ...mon_info, state_dir_path, decrypted_password};
        const mon = _replaceTemplate(mon_template, data);
        const mon_key = `ssh_mon_${mon_info.type}_${mon_info.name}_${mon_info.mon_id}`;
        mons[mon_key] = mon;
    } return mons;
}

function _replaceTemplate(template, data) {
    if (typeof template === "string") {
        const match = template.match(/^__([A-Z0-9_]+)__$/);
        if (match) {
            const dataKey = match[1].toLowerCase();
            return data[dataKey] ?? template; 
        } return template.replace(/__([A-Z0-9_]+)__/g, (_, key) => {
            const dataKey = key.toLowerCase(); return data[dataKey] ?? `__${key}__`;
        });
    }

    if (Array.isArray(template)) return template.map(value => _replaceTemplate(value, data));
    if (template && typeof template === "object") { const result = {};
        for (const [key, value] of Object.entries(template)) result[key] = _replaceTemplate(value, data);
        return result;
    } return template;
}

async function _storeAndgetMonInfos(db, node_id, row) {
    const monInfos = []
    const cpu_usage = row["System CPU Threshold"];
    const ram_usage = row["System RAM Threshold"];
    const disk_usage = row["System Disk Threshold"];
    const folder_rate = row["Folder Growth Threshold"];
    const service_cpu_usage = row["App CPU Threshold"];
    const service_ram_usage = row["App RAM Threshold"];
    const folder_rate_duration = row["Folder Growth Time in Hr(s)"];
    const folder_retention_days = row["Folder Retention (Days)"];
    let services = row["Services"];
    let files = row["Critical Files"];
    let folders = row["Critical Folders"];
    if(cpu_usage || ram_usage || disk_usage) {
        const mon_type = "INFRA"; const serviceNames = services? services.split(";").map(serviceWithPort => 
            serviceWithPort.split(":")[0]).join(" "): undefined;
        if(cpu_usage) monInfos.push(await _addInfraMonToDB(db, node_id, mon_type, "CPU", {cpu_usage}, serviceNames));
        if(ram_usage) monInfos.push(await _addInfraMonToDB(db, node_id, mon_type, "RAM", {ram_usage}, serviceNames));
        if(disk_usage) monInfos.push(await _addInfraMonToDB(db, node_id, mon_type, "DISK", {disk_usage}, serviceNames));
    }
    if (files) {
        files = files.split(";"); for (let fileInfo of files) {
            const file = fileInfo.split(":")[0];
            const file_authorized_users = fileInfo.split(":")[1] || "root";
            const mon = { mon_id: _getUUID(), type: "FILE", status: "healthy", 
                name: "FILE INTEGRITY", file, file_authorized_users
            }; await _storeMonToDB(db, node_id, mon); monInfos.push(mon);
        }
    }
    if(folders && folder_retention_days && folder_rate && folder_rate_duration) {
        folders = folders.split(";");
        for (const folder of folders) {
            const thresholdInfo = { folder_rate };
            let mon = { mon_id: _getUUID(), type: "FOLDER", status: "healthy", 
                name: "FOLDER GROWTH", folder, folder_retention_days, folder_rate_duration
            }; await _storeMonToDB(db, node_id, mon);
            await _storeThresholdToDB(db, node_id, mon.mon_id, thresholdInfo);
            mon={...mon, ...thresholdInfo, duration:2}; monInfos.push(mon);
        }
    }
    if(services) {
        services = services.split(";"); for (let service of services) {
            const [service_name, service_port] = service.split(":");
            const mon = { mon_id: _getUUID(), type: "SERVICE", status: "healthy", 
                name: "STATUS", service_name, service_port
            }; await _storeMonToDB(db, node_id, mon); monInfos.push(mon);
            if(service_cpu_usage || service_ram_usage) {
                const mon_type = "SERVICE INFRA";
                if(service_cpu_usage) monInfos.push(await _addInfraMonToDB(db, node_id, mon_type, 
                    "SERVICE CPU", {service_cpu_usage}, service_name));
                if(service_ram_usage) monInfos.push(await _addInfraMonToDB(db, node_id, mon_type, 
                    "SERVICE RAM", {service_ram_usage}, service_name));
            }
        }
    } return monInfos;
}

async function _addInfraMonToDB(db, node_id, type, name, threshould_info, service_name_or_names) {
    let mon = { mon_id: _getUUID(), type, status: "healthy", name};
    const is_infra_mon = type === "INFRA";
    if(!is_infra_mon) mon.service_name = service_name_or_names;
    await _storeMonToDB(db, node_id, mon);
    await _storeThresholdToDB(db, node_id, mon.mon_id, threshould_info);
    mon = {...mon, ...threshould_info, duration:2};
    if(is_infra_mon) mon.services = service_name_or_names;
    return mon;
}

async function _storeMonToDB(db, node_id, mon_info) {
    const data = { node_id, ...mon_info };
    const columns = Object.keys(data);
    const placeholders = columns.map(() => "?").join(", ");
    const query = `INSERT INTO mons (${columns.join(", ")}) VALUES (${placeholders})`;
    return await dbRunAsync(query, Object.values(data));
}

async function _storeThresholdToDB(db, node_id, mon_id, thresholdInfo) {
    const data = { node_id, mon_id, ...thresholdInfo };
    const columns = Object.keys(data);
    const placeholders = columns.map(() => "?").join(", ");
    const query = `INSERT INTO thresholds (${columns.join(", ")}) VALUES (${placeholders})`;
    return await dbRunAsync(query, Object.values(data));
}

async function _initDB(DB_PATH, DB_CREATION_SQLS) {
    if (!await _createDB(DB_PATH, DB_CREATION_SQLS)) return false;
    if (!await _openDB(DB_PATH)) return false; else return true;
}

async function _createDB(DB_PATH, DB_CREATION_SQLS) {
    try { await accessAsync(DB_PATH, fs.constants.F_OK | fs.constants.W_OK); return true; }
    catch (err) {  // db doesn't exist
        ASBLOG.info("DB doesn't exist, creating and initializing", true);
        try{await mkdirAsync(path.basename(DB_PATH))} catch(err){
            if (err.code != "EEXIST") {ASBLOG.error(`Error creating DB dir, ${err}`, true); return false;}
        } if (!await _openDB(DB_PATH)) return false; // creates the DB file
        
        for (const dbCreationSQL of DB_CREATION_SQLS) try{await dbRunAsync(dbCreationSQL, [])} catch(err) {
            ASBLOG.info(`DB creation DDL failed on: ${dbCreationSQL}, due to ${err}`, true); 
            return false;
        } ASBLOG.info("DB created successfully.", true); return true;    
    }
}

function _openDB(DB_PATH) {
    return new Promise(resolve => {
        if (!db) db = new sqlite3.Database(DB_PATH, sqlite3.OPEN_READWRITE|sqlite3.OPEN_CREATE, err => {
            if (err) {ASBLOG.error(`Error opening DB, ${err}`, true); resolve(false);} 
            else { dbRunAsync = _promisify(db.run.bind(db)); resolve(true); }
        }); else resolve(true);
    });
}

function _promisify(fn) { return (...args) => { return new Promise((resolve, reject) => { 
    fn(...args, function (err, result) { if (err) reject(err); else resolve(result); }); }); }; }

async function _storeNodeToDB(db, nodeInfo) {
    const data = { ...nodeInfo };
    const columns = Object.keys(data);
    const placeholders = columns.map(() => "?").join(", ");
    const query = `INSERT INTO nodes (${columns.join(", ")}) VALUES (${placeholders})`;
    return await dbRunAsync(query, Object.values(data));
}

function _getNodeInfo(row) {
    return { node_id: _getUUID(), ip: row["VM IP"], user: row["SSH User"], port: row["SSH Port"], name: row["Node Name"], 
        status: 'compilant', os_type: row["OS"], password: crypt.encrypt(row["SSH Password"]), pyshell_port: row["Pyshell Port"] }
}

function _getUUID() { return crypto.randomUUID();}