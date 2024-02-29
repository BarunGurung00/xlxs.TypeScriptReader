import * as xlsx from 'xlsx';
import { DynamoDBClient } from "@aws-sdk/client-dynamodb";
import { PutCommand, DynamoDBDocumentClient } from "@aws-sdk/lib-dynamodb";

// Read the Excel file
const WorkBook: xlsx.WorkBook = xlsx.readFile('./dataFolder/general-elections-and-governments.xlsx');

// The third sheet contains the data we want ie the election results from 1918 t0 2019
const sheetName: string =  WorkBook.SheetNames[3];

// Access the sheet object
const dataSheet: xlsx.WorkSheet = WorkBook.Sheets[sheetName];

const range = xlsx.utils.decode_range(dataSheet['!ref'] || ''); 
const num_rows = range.e.r + 1;

console.log("Number of rows with data:", num_rows);

// Convert the sheet data to JSON
const jsonData: any[] = xlsx.utils.sheet_to_json(dataSheet);

// This is the datatype that will be used to store the data in DynamoDB
type DynamoDBData = {
    partyName: string;
    year: number;
    totalVotes: number
}

const allData: DynamoDBData[] = jsonData.map(row => ({
    partyName: row['__EMPTY_1'],
    year: parseInt(row['Election results by party: UK, GB, England, Scotland and Wales']),
    totalVotes: parseInt(row['__EMPTY_3'] || '0', 10)
}));

const region: string ="us-east-1";

const client = new DynamoDBClient({ region });
const documentClient = DynamoDBDocumentClient.from(client);

async function putData() : Promise<void> {
    for (const data of allData) {
        const command = new PutCommand({
            TableName: "Election",
            Item: {
                "partyName": data.partyName,
                "year": data.year,
                "totalVotes": data.totalVotes
            }
        });

        try {
            const response = await documentClient.send(command);
            console.log(response);
        } catch (err:any) {
            console.error("ERROR uploading data Info: " + err.message);
        }
    }
}

putData();
// console.log(allData);