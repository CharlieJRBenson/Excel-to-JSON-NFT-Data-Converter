const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const dir = 'nft_json';

if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir);
}


const workbook = xlsx.readFile('Template For NFT Metadata.xlsx');
const sheet1 = workbook.Sheets[workbook.SheetNames[0]];
const sheet2 = workbook.Sheets[workbook.SheetNames[1]];

const headers = ["NFT #"];
for (let i = 1; ; i++) {
    const cell = sheet1[xlsx.utils.encode_cell({ c: i, r: 7 })];
    if (!cell) break;
    headers.push(cell.v);
}

const artistName = sheet2["A6"].v;
const collectionName = sheet2["B6"].v;
const collectionDescription = sheet2["C6"].v;
const cid = sheet2["D6"].v;


for (let i = 8; ; i++) {
    const nftNumCell = sheet1[xlsx.utils.encode_cell({ c: 0, r: i })];
    if (!nftNumCell) break;
    const nftNum = nftNumCell.v;
    const json = {
        name: `${collectionName} #${nftNum}`,
        description: collectionDescription,
        image: "ipfs://" + cid + "/" + nftNum + ".png",
        date: Date.now(),
        attributes: [],
    };

    for (let j = 0; j < headers.length; j++) {
        json.attributes.push({ trait_type: headers[j], value: sheet1[xlsx.utils.encode_cell({ c: j, r: i })].v });
    }

    json.attributes.push({ trait_type: "artist", value: artistName });
    const filePath = path.join(dir, `${nftNum}.json`);
    fs.writeFileSync(filePath, JSON.stringify(json, null, 2));
}
