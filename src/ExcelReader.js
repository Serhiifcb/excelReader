import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import './ExcelReader.css';

function ExcelReader() {
    const [data, setData] = useState([]);
    
    const printFunc = (data, sumIndex, minIndex, maxIndex) => {
        let mainItem = '';
        let costItem = '';
        data.forEach((el, index) => {
            if (index === sumIndex) {
                for (const key in el) {
                    if (key === 'Стаття витрат') {
                        mainItem = el[key] + ': ';
                    }
                    if (key === 'Відхилення, грн') {
                        mainItem = mainItem + ((el[key] > 0) ? 'Збільшення на ' : 'Зменшення на ') + el[key].toFixed(2) + ' грн. за рахунок:';
                    }
                }
            }
        })
        
        data.forEach((el, index) => {
            let item = '';
            if (index > (minIndex - 1) && index < (maxIndex + 1)) {
                for (const key in el) {
                    if (key === 'Стаття витрат') {
                        item = item + el[key];
                    }
                    if (key === 'Відхилення за рахунок ціни, грн') {
                        if (el[key] !== 0) {
                            costItem = costItem + ((el[key].toFixed(2) > 0) 
                            ? (' збільшення на ' + el[key].toFixed(2) + ` грн. внаслідок зростання ціни${(minIndex === sumIndex) ? '' : (' на ' + item)};`) 
                            : (el[key].toFixed(2) < 0) 
                            ? (' зменшення на ' + el[key].toFixed(2) + ` грн. внаслідок зниження ціни${(minIndex === sumIndex) ? '' : (' на ' + item)};`)
                            : '');
                        }
                    }
                    if (key === 'Відхилення за рахунок норми, грн') {
                        if (el[key] !== 0) {
                            costItem = costItem + ((el[key].toFixed(2) > 0) 
                            ? (' збільшення на ' + el[key].toFixed(2) + ` грн. внаслідок зростання норми${(minIndex === sumIndex) ? '' : (' на ' + item)};`) 
                            : (el[key].toFixed(2) > 0)
                            ? (' зменшення на ' + el[key].toFixed(2) + ` грн. внаслідок зниження норми${(minIndex === sumIndex) ? '' : (' на ' + item)};`)
                            : '');
                        }
                    }
                }
            }
        });
        return (mainItem + costItem);
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
            // виведення даних в консоль
            console.log(data);
        };
        reader.readAsBinaryString(file);
    }

    return (
        <div>
            <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
            {/* Відображення даних, якщо потрібно */}
            {/* <pre>{JSON.stringify(data, null, 2)}</pre> */}
            <div className='result'>
                {printFunc(data, 1, 1, 1)} <br/>
                {printFunc(data, 23, 2, 22)} <br/>
                {printFunc(data, 26, 24, 25)} <br/>
                {printFunc(data, 35, 29, 34)} <br/>
                {printFunc(data, 38, 36, 37)} <br/>
            </div>
        </div>
    );
}

export default ExcelReader;