const { extractData } = require('D:/Projetos/ExportExcelApps/app.js');
const schedule = require('node-schedule');

const data = new Date();
let ano = data.getFullYear().toString();
let mes = data.toLocaleString('pt-br', { month: 'short' }).replace('.', '');

const extractArqs = [
    {
        id: 'idObjeto',
        filterYear: ano,
        filterMonth: mes,
        archiveName: 'NomeDoArquivo.csv',
        archivePath: 'Diretorio onde os arquivos serao salvos',
        order: ['Array para ordenar as colunas corretamente']
    },
    {
        id: 'idObjeto2',
        filterYear: ano,
        filterMonth: mes,
        archiveName: 'NomeDoArquivo2.csv',
        archivePath: 'Diretorio onde os arquivos serao salvos',
        order: ['Array para ordenar as colunas corretamente']
    }
];

const job = schedule.scheduleJob('0 40 23 * * *', () => {
    registraLogs('Iniciando processo de extração');
    async function processaDados(dados) {
        for (const extracao of dados) {
            registraLogs('Extracting: ' + extracao.archiveName);
            await extractData(extracao.id, extracao.filterYear, extracao.filterMonth, extracao.archiveName, extracao.archivePath, extracao.order);
        }
    }

    processaDados(extractArqs).then(() => {
        registraLogs('Processo concluído');
    });

});