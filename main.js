const { google } = require('googleapis');
const keys = require('./keys.json');

const client = new google.auth.JWT(
    keys.client_email, 
    null, 
    keys.private_key, 
    ['https://www.googleapis.com/auth/spreadsheets']
);

client.authorize(function(err, tokens){

    if(err){
        console.log(err);
        return;
    } else {
        console.log('Connected to the Spreadsheet!');
        gsrun(client);
    }

});

async function gsrun(cl){

    const gsapi = google.sheets({version:'v4', auth: cl });

    const opt = {
        spreadsheetId: '1InUcO6QBF2PZ2jUTub8zaqVOVbWSo2mr_emeadSGcYQ',
        range: 'engenharia_de_software!A4:F'
    };

    let data = await gsapi.spreadsheets.values.get(opt);
    console.log('Getting data from the sheet.');
    let array = data.data.values;

    console.log('Calculating the average of students.')
    media = [];
    for(let n = 0; n<24; n++){
        media[n] = (parseInt((Number(array[n][3]) + Number(array[n][4]) + Number(array[n][5]))/3));
    };
    console.log(media);

    console.log('Getting test results.')
    let again = 'Reprovado por Falta';
    let max = 60/4;
    situation = [];
    for(let n = 0; n<24; n++){
        if(max<Number(array[n][2])){
            situation[n] = again;
        } else if(media[n] < 50){
            situation[n] = 'Reprovado por Nota';
        }else if(50 <= media[n] && media[n] < 70){
            situation[n] = 'Exame Final';
        }else if(media[n] >= 70){
            situation[n] = 'Aprovado';
        }
    }
    console.log(situation);

    console.log('Calculating final grades.')
    naf = [];
    for(let n = 0; n<24; n++){
        if(situation[n] == 'Exame Final'){
            naf[n] = 100 - media[n];
        }else if(situation[n] != 'Exame Final'){
            naf[n] = '0';
        }
    }


    console.log('Results obtained.')
    let final_values = []; 
    for(n=0;n<24;n++){
        final_values[n] = [];
    };
    for (let n = 0; n < 24; n++){
        final_values[n][0] = situation[n];
        final_values[n][1] = naf[n];
    };
    console.log(final_values);

   console.log('Uploading exam data.')
    const upt = {
    spreadsheetId: '1InUcO6QBF2PZ2jUTub8zaqVOVbWSo2mr_emeadSGcYQ',
    range: 'engenharia_de_software!G4:H',
    valueInputOption: 'USER_ENTERED',
    resource: { values: final_values}
};

    let res = await gsapi.spreadsheets.values.update(upt);
    

}