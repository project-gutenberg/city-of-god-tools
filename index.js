const fs = require('fs-extra');
const path = require('path');
const humps = require('humps');
const xlsx = require('xlsx');
const po = require('gettext-parser').po;

if (!String.prototype.padStart)
    String.prototype.padStart = function padStart(targetLength, padString) {
        targetLength = targetLength >> 0; //floor if number or convert non-number to 0;
        padString = String(padString || ' ');
        if (this.length > targetLength)
            return String(this);
        else {
            targetLength = targetLength - this.length;
            if (targetLength > padString.length)
                padString += padString.repeat(targetLength / padString.length); //append to original to ensure we are longer than needed
            return padString.slice(0, targetLength) + String(this);
        }
    };

const workbook = xlsx.readFile(path.join(__dirname, '翻译内容 - 精简.xlsx'));

const generatePo = () => {
    workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        const data = {};

        Object.keys(sheet).forEach(key => {
            if (key.startsWith('!'))
                return;

            const columnIndex = key.substr(0, 1);
            const rowIndex = parseInt(key.substr(1), 10);
            if (isNaN(rowIndex))
                throw new Error(`Unable to parse ${key}.`);
            if (rowIndex === 1)
                return;

            data[rowIndex] = data[rowIndex] || {};
            switch (columnIndex) {
                case 'A':
                    data[rowIndex].key = sheet[key].v;
                    break;
                case 'B':
                    data[rowIndex].cn = sheet[key].v;
                    break;
                case 'C':
                    data[rowIndex].en = sheet[key].v;
                    break;
                case 'D':
                    data[rowIndex].desc = sheet[key].v;
                    break;
            }
        });

        const keys = Object.keys(data);

        if (sheetName === 'Msg') {
            const re = /^[^"]+?"(.+?)"/i;
            keys.forEach(key => {
                const entry = data[key];
                let match = re.exec(entry.cn.trim());
                if (match)
                    entry.cn = match[1];
                if (entry.en) {
                    match = re.exec(entry.en.trim());
                    if (match)
                        entry.en = match[1];
                }
            });
        }

        const SIZE = 150;
        let pageSize = SIZE;
        let page = Math.floor(keys.length / SIZE);
        if (page === 0) {
            page = 1;
            pageSize = keys.length;
        } else {
            const mod = keys.length % SIZE;
            if (mod >= SIZE * 0.5)
                page += 1;
            else
                pageSize += Math.ceil(mod / page);
        }

        // po
        for (let i = 0; i < page; i++) {
            const number = (i + 1).toString().padStart(2, '0');
            const srcFileName = path.join(__dirname, 'output', 'pot', `${humps.pascalize(sheetName)}_${number}.pot`);
            const transFileName = path.join(__dirname, 'output', 'en-US', `${humps.pascalize(sheetName)}_${number}.po`);
            const srcOutput = {
                charset: "utf-8",
                headers: {
                    "content-type": "text/plain; charset=utf-8",
                    "plural-forms": "nplurals=2; plural=(n!=1);"
                },
                translations: {}
            };
            const transOutput = Object.assign({}, srcOutput);

            for (let j = i * pageSize; j <= (i + 1) * pageSize - 1; j++) {
                if (j >= keys.length)
                    break;

                const entry = data[keys[j]];
                if (!entry.key || entry.key === '' || !entry.cn || entry.cn === '')
                    continue;

                const srcObj = {
                    msgctxt: entry.key,
                    msgid: entry.cn,
                    msgstr: ''
                };
                if (entry.desc)
                    srcObj.comments = {
                        translator: entry.desc
                    };
                const transObj = entry.en ? Object.assign({}, srcObj, {msgstr: entry.en}) : null;

                srcOutput.translations[entry.key] = srcOutput.translations[entry.key] || {};
                srcOutput.translations[entry.key][entry.cn] = srcObj;
                if (transObj) {
                    transOutput.translations[entry.key] = transOutput.translations[entry.key] || {};
                    transOutput.translations[entry.key][entry.cn] = transObj;
                }
            }

            fs.outputFileSync(srcFileName, po.compile(srcOutput));
            fs.outputFileSync(transFileName, po.compile(transOutput));
        }
    });
};

const generateXlsx = () => {
    const srcPath = path.join(__dirname, 'en-US');
    const xlsxFileName = path.join(__dirname, 'output', '翻译导出20170828.xlsx');
    // const refs = {
    //     loading: 'A1:I207'
    // };

    workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        // sheet['!ref'] = refs[sheetName] || sheet['!ref'];

        const transData = {};
        fs.readdirSync(srcPath).map(f => path.join(srcPath, f)).filter(f => fs.statSync(f).isFile()).forEach(poFile => {
            if (!path.basename(poFile).toLowerCase().startsWith(`${sheetName.toLowerCase()}_`))
                return;

            let translations = po.parse(fs.readFileSync(poFile)).translations;
            if (!translations)
                return;
            delete translations[''];

            Object.keys(translations).forEach(id => {
                let translation = translations[id];
                translation = translation[Object.keys(translation)[0]];
                if (!translation.msgstr || !Array.isArray(translation.msgstr) || translation.msgstr.length < 1)
                    return;
                if (translation.msgstr[0] !== '')
                    transData[id] = translation.msgstr[0];
            });
        });
        //console.log(JSON.stringify(transData, null, 2));

        Object.keys(sheet).forEach(key => {
            const columnIndex = key.substr(0, 1);
            if (columnIndex !== 'A')
                return;
            const rowIndex = parseInt(key.substr(1), 10);
            if (rowIndex === 1 || isNaN(rowIndex))
                return;

            const id = sheet[key].v;
            if (!transData[id])
                return;

            sheet[`C${rowIndex}`] = sheet[`C${rowIndex}`] || {t: 's'};
            if (sheetName === 'Msg') {
                const en = sheet[`B${rowIndex}`].v;
                if (en) {
                    const index = en.indexOf('=');
                    sheet[`C${rowIndex}`].v = `${en.substr(0, index)}= "${transData[id].replace(/"/g, '\\"')}"`;
                }
            } else
                sheet[`C${rowIndex}`].v = transData[id];

            console.log(key, id, transData[id]);
        });
        console.log(sheetName, sheet['!ref']);
        console.log();
    });

    xlsx.writeFile(workbook, xlsxFileName);
};

generatePo();
