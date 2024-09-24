const input_file = document.querySelector('#input-file');
const output_spreadsheet = document.querySelector('#output-spreadsheet');
const containerInputFile = document.querySelector('#container-input-file');
const downloadSpreadsheet = document.querySelector('#download-spreadsheet');

containerInputFile.addEventListener('click', () => {
    input_file.click();
});

const db = [];
const finalArray = []; 
const dataFromAllCustomers = [['CLIENTE', 'GRUPO', 'CPF/CNPJ', 'EMAIL', 'TELEFONE']]; // Cabeçalhos

const createSpreadsheet = () => {
    output_spreadsheet.innerHTML = ''; // Limpa o conteúdo anterior

    if (dataFromAllCustomers.length > 0) {
        const table = document.createElement('table');
        table.classList.add('table');
        output_spreadsheet.appendChild(table);

        // Adiciona cabeçalhos
        const headerRow = document.createElement('tr');
        table.appendChild(headerRow);

        dataFromAllCustomers[0].forEach(header => {
            const th = document.createElement('th');
            th.textContent = header; // Define o texto do cabeçalho
            headerRow.appendChild(th);
        });

        // Adiciona os dados das linhas
        for (let i = 1; i < dataFromAllCustomers.length; i++) { // Começa do índice 1 para pular os cabeçalhos
            const tr = document.createElement('tr');
            tr.classList.add('tr');
            table.appendChild(tr);

            dataFromAllCustomers[i].forEach(value => {
                const td = document.createElement('td');
                td.classList.add('td');
                td.textContent = value; // Define o texto da célula
                tr.appendChild(td);
            });
        }
    } else {
        alert('Arquivo em Branco ou corrompido');
    }
}

downloadSpreadsheet.addEventListener('click', () => {
    if (dataFromAllCustomers.length > 1) { 
        downloadSpreadsheet.style.backgroundColor = 'rgb(76, 175, 80)'; 
        downloadSpreadsheet.textContent = 'Baixando arquivo'

        const wb = XLSX.utils.table_to_book(document.querySelector('.table'), { sheet: "Sheet1" });
        XLSX.writeFile(wb, 'planilhas.xlsx');

        setTimeout(() => {
            downloadSpreadsheet.style.transition = '500ms'
            downloadSpreadsheet.style.backgroundColor = 'rgb(5, 77, 119)';
            downloadSpreadsheet.textContent = 'Baixar arquivo .xlsx'
        }, 1000)
    } else {
        downloadSpreadsheet.style.backgroundColor = 'rgb(255, 64, 64)';
        downloadSpreadsheet.textContent = 'Abra um arquivo antes de baixar'

        setTimeout(() => {
            downloadSpreadsheet.style.transition = '500ms'
            downloadSpreadsheet.style.backgroundColor = 'rgb(5, 77, 119)';
            downloadSpreadsheet.textContent = 'Baixar arquivo .xlsx'
        }, 1500)
    }
});


const spreadsheetOrganizer = () => {
    const organizedSpreadsheet = [];
    const indexesImportantFields = [];

    const separateByRelevance = () => {
        indexesImportantFields.forEach((rowIndex) => {
            const firstAddition = rowIndex + 1; 
            const arraysConcat = [db[firstAddition]];
            organizedSpreadsheet.push(arraysConcat);
        });   

        finalArray.length = 0;
        let pairBuffer = []; 

        organizedSpreadsheet.forEach((rows) => {
            rows.forEach((row) => {
                const filteredObject = {};
                for (const [key, value] of Object.entries(row)) {
                    if (value !== '') {
                        filteredObject[key] = value;
                    }
                }

                if (Object.keys(filteredObject).length > 0) {
                    pairBuffer.push(filteredObject); 
                }

                if (pairBuffer.length === 2) {
                    finalArray.push([...pairBuffer]); 
                    pairBuffer = []; 
                }
            });
        });

        const regex = [
            /^(?!.*\d)([A-ZÀ-Ü]+(?:\s+[A-ZÀ-Ü]+)+)$/,  // Nomes
            /^0\d{5}$/,  // Números (6 dígitos, começando com 0)
            /^\d{3}\.\d{3}\.\d{3}-\d{2}$/,   // CPF
            /^\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}$/,   // CNPJ
            /Telefone:/, // String Telefone
            /^\d{2}$/, // DDD
            /^\d{8}$/, // Telefone 8 dígitos
            /^9\d{8}$/, // Telefone 9 dígitos com "9"
            /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/   // Emails
        ];

        finalArray.forEach((row) => {
            let dddFound = null; // DDD encontrado
            let telefoneFound = null; // Telefone encontrado
            const customerData = [];

            row.forEach((rowTwo) => {
                for (const [key, value] of Object.entries(rowTwo)) {
                    if (regex[4].test(value)) { // Se for Telefone
                        const nextIndex = Object.keys(rowTwo).indexOf(key) + 1; // Próximo índice
                        let collectedValues = []; // Array para armazenar os valores coletados

                        for (let i = 0; i < 2; i++) { // Coletar os próximos dois valores
                            if (nextIndex + i < Object.keys(rowTwo).length) {
                                collectedValues.push(rowTwo[Object.keys(rowTwo)[nextIndex + i]]);
                            }
                        }

                        collectedValues.forEach((val) => {
                            if (regex[5].test(val)) {
                                dddFound = val; // DDD encontrado
                            }
                            if (regex[6].test(val) || regex[7].test(val)) {
                                telefoneFound = val; // Telefone encontrado
                            }
                        });
                    }

                    if (typeof value === 'string' && regex[0].test(value)) {
                        customerData.push(value); // Nome
                    }
                    
                    if (typeof value === 'string' && regex[1].test(value)) {
                        customerData.push(value); // Grupo
                    }

                    if (typeof value === 'string' && regex[2].test(value)) {
                        customerData.push(value); // CPF
                    } else if (typeof value === 'string' && regex[3].test(value)) {
                        customerData.push(value); // CNPJ
                    } 

                    if (typeof value === 'string' && regex[8].test(value)) {
                        customerData.push(value); // Email
                    }
                }
            });

            // Adicionar DDD padrão se não encontrado
            const dddPadrao = 'DDD não informado'; 
            if (!dddFound && telefoneFound) {
                dddFound = dddPadrao; // DDD padrão
            }

            if (dddFound && telefoneFound) {
                customerData.push(`(${dddFound}) ${telefoneFound}`); // Adiciona telefone completo
            }

            dataFromAllCustomers.push(customerData); // Adiciona dados do cliente
        });
        createSpreadsheet(); // Gera a planilha após coletar todos os dados
    };

    db.forEach((row, index) => {
        for (const [key, value] of Object.entries(row)) {
            if (value.trim() === 'G49') { 
                indexesImportantFields.push(index);
            }
        }
    });

    separateByRelevance();
}

input_file.addEventListener('change', event => {
    const file = event.target.files[0]; 
    const reader = new FileReader();

    reader.onload = (event) => {
        const data = new Uint8Array(event.target.result); 
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0]; 
        const sheet = workbook.Sheets[sheetName]; 
        const json = XLSX.utils.sheet_to_json(sheet, {
            raw: false,
            dateNF: 'dd/mm/yy',
            defval: '',
        }); 
        db.length = 0; 
        db.push(...json);
        spreadsheetOrganizer(); // Organiza a planilha após carregar os dados
    }

    reader.readAsArrayBuffer(file); 
});
