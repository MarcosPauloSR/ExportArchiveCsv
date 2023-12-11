const enigma = require('enigma.js');
const WebSocket = require('ws');
const path = require('path');
const fs = require('fs');
const schema = require('enigma.js/schemas/12.170.2.json');
const createCsvWriter = require('csv-writer').createObjectCsvWriter;

// Configurações existentes
const engineHost = 'host do Qlik Sense';
const enginePort = 4747;
const appId = 'Id do aplicativo';
const userDirectory = 'Diretorio de Usuarios';
const userId = 'Usuario';
const certificatesPath = 'Diretorio dos certificados';

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

  let archive = archivePath + archiveName;

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

          registraLogs(`Total rows in hypercube: ${totalRows}`);
          registraLogs(`Total columns in hypercube: ${totalColumns}`);

          let data = [];

          const fetchBatchData = (qTop, qHeight) => {
            registraLogs(`Fetching data from row ${qTop} for ${qHeight} rows`);

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
            //console.log(`Total rows fetched: ${data.length}`);
            registraLogs(`Total rows fetched: ${data.length}`);

            data = data.filter(row => {
              const year = row[correctOrder.indexOf('Ano')];
              const month = row[correctOrder.indexOf('Mês')];
              return year === filterYear.toString() && month === filterMonth;
            })

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
                registraLogs('Dados exportados com sucesso para ' + archive);
              })
              .catch(err => {
                registraLogs('Erro ao escrever no arquivo CSV: ' + err);
                session.close();
              });
          });
        });
      });
    }).catch(err => {
      registraLogs('Erro ao recuperar dados do Qlik Sense: ' + err);
      session.close();
      throw err;
    });
  }).catch(err => {
    registraLogs('Erro ao abrir a sessão: ' + err);
    session.close();
    throw err;
  });
}

function getFormattedDate() {
  const date = new Date();
  const year = date.getFullYear();
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const day = date.getDate().toString().padStart(2, '0');
  return `${day}-${month}-${year}`;
}

function registraLogs(message) {
  const timestamp = new Date().toLocaleString();
  const logMessage = `${timestamp} - ${message}\n`;

  // Nome do arquivo baseado na data atual
  const logFileName = `log_${getFormattedDate()}.txt`;
  const logFilePath = path.join('D:/Projetos/ExportExcelApps/logs/', logFileName);

  console.log(logMessage);

  fs.appendFileSync(logFilePath, logMessage, 'utf8');

  if (message.toUpperCase().match('ERRO')) {
    sendLogMail(logFilePath, getFormattedDate());
  }
}

module.exports.extractData = extractData;
module.exports.registraLogs = registraLogs