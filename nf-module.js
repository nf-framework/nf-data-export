import path from 'path';
import fs from 'fs';
import mime from 'mime';
import { registerLibDir, registerCustomElementsDir } from "@nfjs/front-server";
import { web } from '@nfjs/back';

const meta = {
    require: {
        after: '@nfjs/back'
    }
};

const __dirname = path.join(path.dirname(decodeURI(new URL(import.meta.url).pathname))).replace(/^\\([A-Z]:\\)/, "$1");

async function init() {
    registerLibDir('xlsx-populate', null, { denyPathReplace: true, minify: 'deny' });
    registerCustomElementsDir('@nfjs/data-export/components');

    web.on('GET', '/@nfjs/data-export/getReportTemplate', (context) => {
        const file = `${__dirname}/reportTemplate.xlsx`;

        const mimetype = mime.getType(file);
        context.type('Content-Type', mimetype);

        const headers = {
            'Content-Disposition': `attachment; filename=${encodeURIComponent(`reportTemplate.xlsx`)}`,
            'Content-Transfer-Encoding': 'binary'
        }
        context.headers(headers);

        let filestream = fs.createReadStream(file);
        context.send(filestream);
    });
}

export {
    meta,
    init
};
