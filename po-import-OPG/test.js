const axios = require("axios");

const fromDate = "20210401"; // YYYYMMDD
const toDate = "20230430";

const xmlRequest = `
<ENVELOPE>
    <HEADER>
        <VERSION>1</VERSION>
        <TALLYREQUEST>Export</TALLYREQUEST>
        <TYPE>Data</TYPE>
        <ID>Day Book</ID>
    </HEADER>

    <BODY>
        <DESC>
            <STATICVARIABLES>
                <SVFROMDATE TYPE="Date">${fromDate}</SVFROMDATE>
                <SVTODATE TYPE="Date">${toDate}</SVTODATE>
                <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
            </STATICVARIABLES>
        </DESC>
    </BODY>
</ENVELOPE>
`;

axios.post("http://localhost:9321", xmlRequest, {
    headers: {
        "Content-Type": "application/xml"
    }
})
.then(res => {
    console.log(res.data);   // XML response
})
.catch(console.error);