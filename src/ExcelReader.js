import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';

function ExcelReader() {
    const [data, setData] = useState([]);
    const [wastepaper, setWastepaper] = useState('');
    
    const paperFunc = (data, sumIndex, minIndex, maxIndex) => {
        let costItem = '';
        data.forEach((el, index) => {
            if (index === sumIndex) {
                for (const key in el) {
                    if (key === 'Стаття витрат') {
                        costItem = el[key] + ': ';
                    }
                    if (key === 'Відхилення, грн') {
                        costItem = costItem + ((el[key] > 0) ? 'Збільшення на ' : 'Зменшення на ') + el[key].toFixed(2) + ' за рахунок: ';
                    }
                }
            }
        })
        
        data.forEach((el, index) => {
            let item = '';
            if (index > (minIndex - 1) && index < (maxIndex +1)) {
                for (const key in el) {
                    if (key === 'Стаття витрат') {
                        item = item + el[key];
                    }
                    if (key === 'Відхилення за рахунок ціни, грн') {
                        if (el[key] !== 0) {
                            costItem = costItem + (el[key] > 0) 
                            ? ('збільшення на ' + el[key].toFixed(2) + ` за рахунок зростання ціни на ${item}; `) 
                            : ('зменшення на ' + el[key].toFixed(2) + ` за рахунок зниження ціни на ${item}; `);
                        }
                    }
                    if (key === 'Відхилення за рахунок норми, грн') {
                        if (el[key] !== 0) {
                            costItem = costItem + (el[key] > 0) 
                            ? ('збільшення на ' + el[key].toFixed(2) + ` за рахунок зростання норми на ${item}; `) 
                            : ('зменшення на ' + el[key].toFixed(2) + ` за рахунок зниження норми на ${item}; `);
                        }
                    }
                }
            }
        });
        return console.log(costItem);
    }

    const chemistryFunc = (data) => {
        let chemistry = 'Хімія: ';
        data.forEach((el, index) => {
            if (index > 1 && index < 24) {
                for (const key in el) {
                    if (key === 'Відхилення, грн' && index === 23) {
                        chemistry = chemistry + ((el[key] > 0) ? 'Збільшення на ' : 'Зменшення на ') + el[key].toFixed(2) + ' за рахунок: ';
                    }
                    if (key === 'Відхилення за рахунок ціни, грн') {
                        if (el[key] !== 0) {
                            chemistry = chemistry + ((el[key] > 0) 
                            ? ('збільшення на ' + el[key].toFixed(2) + ' за рахунок зростання ціни; ') 
                            : ('зменшення на ') + el[key].toFixed(2) + ' за рахунок зниження ціни; ');
                        }
                    }
                    if (key === 'Відхилення за рахунок норми, грн') {
                        if (el[key] !== 0) {
                            chemistry = chemistry + ((el[key] > 0) 
                            ? ('збільшення на ' + el[key].toFixed(2) + ' за рахунок зростання норми; ') 
                            : ('зменшення на ') + el[key].toFixed(2) + ' за рахунок зниження норми; ');
                        }
                    }
                }
            }
        });
        return console.log(chemistry);
    }
    
    const handleFileUpload = e => {
        const file = e.target.files[0];
        const reader = new FileReader();
        reader.onload = (evt) => {
            // парсинг даних
            const bstr = evt.target.result;
            const wb = XLSX.read(bstr, { type: 'binary' });
            // отримання першого листа
            const wsname = wb.SheetNames[0];
            const ws = wb.Sheets[wsname];
            // конвертування в масив JSON
            const data = XLSX.utils.sheet_to_json(ws);
            setData(data);
            paperFunc(data, 1, 1, 1);
            // виведення даних в консоль
            console.log(data);
        };
        reader.readAsBinaryString(file);
    }

    return (
        <div>
            <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
            {/* Відображення даних, якщо потрібно */}
            <pre>{JSON.stringify(data, null, 2)}</pre>
        </div>
    );
}

export default ExcelReader;