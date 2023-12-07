const { extractData } = require('D:/Projetos/ExportExcelApps/app.js');
const schedule = require('node-schedule');

const data = new Date();
let ano = data.getFullYear().toString();
let mes = data.toLocaleString('pt-br', { month: 'short' }).replace('.', '');

const extractArqs = [
    {
        id: 'jAZqVF',
        filterYear: ano,
        filterMonth: mes,
        archiveName: 'NPS_Loja_' + mes + '_' + ano + '.csv',
        archivePath: 'D:/QlikSense/AppsData/Resources/ArquivosNPS/',
        order: [
            'Ano', 'Mês', 'Filial', 'Cnpj', 'Departamento',
            'datareferenciainicial', 'datareferenciafinal', 'dataintegracao',
            'Status', 'NPS', 'Sucesso%'
        ]
    },
    {
        id: 'Ammd',
        filterYear: ano,
        filterMonth: mes,
        archiveName: 'NPS_Consultor_' + mes + '_' + ano + '.csv',
        archivePath: 'D:/QlikSense/AppsData/Resources/ArquivosNPS/',
        order: [
            'Ano', 'Mês', 'Consultor', 'Cnpj', 'Departamento',
            'datareferenciainicial', 'datareferenciafinal', 'dataintegracao',
            'Status', 'NPS', 'Sucesso%'
        ]
    },
    {
        id: 'MAARFrf',
        filterYear: ano,
        filterMonth: mes,
        archiveName: 'NPS_Vendedor_' + mes + '_' + ano + '.csv',
        archivePath: 'D:/QlikSense/AppsData/Resources/ArquivosNPS/',
        order: [
            'Ano', 'Mês', 'Vendedor', 'Cnpj', 'Departamento',
            'datareferenciainicial', 'datareferenciafinal', 'dataintegracao',
            'Status', 'NPS', 'Sucesso%'
        ]
    }
];

const job = schedule.scheduleJob('0 40 23 * * *', () => {
    console.log('Iniciando processo de extração')
    async function processaDados(dados) {
        for (const extracao of dados) {
            console.log('Extracting: ', extracao.archiveName);
            await extractData(extracao.id, extracao.filterYear, extracao.filterMonth, extracao.archiveName, extracao.archivePath, extracao.order);
        }
    }

    processaDados(extractArqs).then(() => {
        console.log('Processo concluído');
    });

})