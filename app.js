const enigma = require('enigma.js');
const WebSocket = require('ws');
const path = require('path');
const fs = require('fs');
const schema = require('enigma.js/schemas/12.170.2.json');
const createCsvWriter = require('csv-writer').createObjectCsvWriter;

// Configurações existentes
const engineHost = 'awswqlik03.sagagyn.local';
const enginePort = 4747;
const appId = 'c74fb7f6-47c8-4a0a-9cd3-2b5e518887e7';
const userDirectory = 'SAGAGYN';
const userId = 'qlikview';
const certificatesPath = 'C:/ProgramData/Qlik/Sense/Repository/Exported Certificates/CerticadosQlik/';

/*const filterYear = '2023';
const filterMonth = 'nov';*/

//let archive = 'D:/QlikSense/AppsData/Resources/ArquivosNPS/NPS_Loja_' + filterMonth + '_' + filterYear + '.csv';

const readCert = (filename) => fs.readFileSync(path.resolve(certificatesPath, filename));


const session = enigma.create({
  schema,
  url: `wss://${engineHost}:${enginePort}/sense/app/${appId}`,
  createSocket: (url) => new WebSocket(url, {
    ca: [readCert('root.pem')],
    key: readCert('client_key.pem'),
    cert: readCert('client.pem'),
    headers: {
      'X-Qlik-User': `UserDirectory=${encodeURIComponent(userDirectory)}; UserId=${encodeURIComponent(userId)}`,
    },
  }),
});

async function extractData(idObject, filterYear, filterMonth, archiveName, archivePath, order) {

  let archive = archivePath+archiveName;

  return session.open().then((global) => {
    console.log('Session was opened successfully');
    const correctOrder = order;

    const csvWriter = createCsvWriter({
      path: archive,
      header: correctOrder.map(header => ({ id: header, title: header })),
      fieldDelimiter: ';',
      append: false,
      encoding: 'utf-8'
    });

    return global.openDoc(appId).then(app => {
      return app.getObject(idObject).then(object => {
        return object.getLayout().then(layout => {
          const totalRows = layout.qHyperCube.qSize.qcy;
          const totalColumns = layout.qHyperCube.qSize.qcx;

          console.log(`Total rows in hypercube: ${totalRows}`);
          console.log(`Total columns in hypercube: ${totalColumns}`);

          let data = [];

          const fetchBatchData = (qTop, qHeight) => {
            console.log(`Fetching data from row ${qTop} for ${qHeight} rows`);

            return object.getHyperCubeData('/qHyperCubeDef', [{
              qTop: qTop, qLeft: 0, qWidth: totalColumns, qHeight: qHeight
            }]).then(pages => {
              const batchData = pages[0].qMatrix.map(row => row.map(cell => cell.qText));
              data = batchData.concat(data);

              if (qTop > 0) {
                const nextQTop = Math.max(qTop - qHeight, 0);
                const nextQHeight = Math.min(qHeight, qTop);
                return fetchBatchData(nextQTop, nextQHeight);
              }

              return data;
            });
          };

          const batchHeight = 700; // Ajuste conforme necessário
          let initialQTop = Math.max(totalRows - batchHeight, 0);
          const initialQHeight = Math.min(batchHeight, totalRows);

          if (initialQTop + initialQHeight > totalRows) {
            // Ajusta qTop para garantir que não exceda o total de linhas
            initialQTop = totalRows - initialQHeight;
          }

          return fetchBatchData(initialQTop, initialQHeight).then(data => {
            console.log(`Total rows fetched: ${data.length}`);

            data = data.filter(row => {
              const year = row[correctOrder.indexOf('Ano')];
              const month = row[correctOrder.indexOf('Mês')];
              return year === filterYear.toString() && month === filterMonth;
            })

            //fs.writeFileSync('exported_table.csv', '\ufeff', { encoding: 'utf8', flag: 'w' });
            fs.writeFileSync(archive, '\ufeff', { encoding: 'utf8', flag: 'w' });

            return csvWriter
              .writeRecords(
                data.map(row => {
                  return correctOrder.reduce((obj, header, index) => {
                    obj[header] = row[index];
                    return obj;
                  }, {});
                })
              )
              .then(() => {
                console.log('Dados exportados com sucesso para ' + archive);
                //session.close();
              })
              .catch(err => {
                console.error('Erro ao escrever no arquivo CSV', err);
                session.close();
              });
          });
        });
      });
    }).catch(err => {
      console.error('Erro ao recuperar dados do Qlik Sense', err);
      session.close();
      throw err;
    });
  }).catch(err => {
    console.error('Erro ao abrir a sessão', err);
    session.close();
    throw err;
  });
}

module.exports.extractData = extractData;