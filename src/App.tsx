import React, {ChangeEvent, useRef, useState} from 'react';
import {read, utils, writeFile} from "xlsx";
import moment, {Moment} from "moment";

class Person {
  name: string = '';
  id: string = '';

  static fromString = (s: string): Person => {
    const [name, id] = s.split("_");
    return {name, id} as Person;
  }

  static toStr = (p: Person) => `${p.name}_${p.id}`;
}


const App: React.FC = () => {
  const [months, setMonths] = useState<Moment[]>([]);
  const [data, setData] = useState<Map<string, Moment[]>>(new Map());
  const [persons, setPersons] = useState<Person[]>([]);
  const [fileName, setFileName] = useState("");

  const personInputRef = useRef<any>(null);
  const excelInputRef = useRef<any>(null);
  const tableRef = useRef<any>(null);


  const readStringAsYearMonth = (s: string) => moment(s, 'YYYYMM');
  const yearMonthToDateString = (yearMonth: Moment, isStart: boolean) => {
    const lastYM = moment().subtract(1, "month");
    return yearMonth.year() === lastYM.year() && yearMonth.month() === lastYM.month() && !isStart ?
      "/" :
      yearMonth.format('YYYY/MM/DD')
  };
  const yearMonthToString = (yearMonth: Moment) => yearMonth.format('YYYYMM');

  const sortYearMonth = (yearMonths: Moment[]) => {
    const sort = yearMonths.sort((a, b) => a.diff(b));
    return [sort[0], sort[sort.length - 1]]
  }

  const loadExcel = (event: ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (files) {
      setFileName("[汇总]" + files[0].name);
      let fr = new FileReader();
      fr.readAsBinaryString(files[0]);
      fr.onload = (f) => {

        const data = new Map<string, Moment[]>();
        const months: Moment[] = [];

        const res = f.target?.result as string;
        const workBook = read(res, {type: "binary"});
        workBook.SheetNames.forEach(sheetName => {
          const ds = utils.sheet_to_json<Person>(workBook.Sheets[sheetName], {header: ['name', 'id']});
          const month = readStringAsYearMonth(sheetName);
          months.push(month);
          ds.splice(0, 3);
          ds.forEach(row => {
            row.name = row.name.trim();
            row.id = row.id.trim();
            const key = Person.toStr(row);
            if (!data.has(key)) {
              data.set(key, [month]);
            } else {
              data.get(key)?.push(month);
            }
          })
        })

        setMonths(months);
        setData(data);
      };
    }
  }

  const loadPerson = (event: ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (files) {
      let fr = new FileReader();
      fr.readAsBinaryString(files[0]);
      fr.onload = (f) => {
        const res = f.target?.result as string;
        const workBook = read(res, {type: "binary"});
        const ds = utils.sheet_to_json<Person>(
          workBook.Sheets[workBook.SheetNames[0]],
          {header: ['name', 'id']}
        );
        ds.splice(0, 1);
        ds.forEach(p => {
          p.name = p.name.trim();
          p.id = p.id.trim();
        });
        setPersons(ds);
      };
    }
  }

  const resetPerson = () => {
    setPersons([]);
    personInputRef.current.value = null;
  };

  const resetExcel = () => {
    setData(new Map());
    setMonths([]);
    setFileName("");
    excelInputRef.current.value = null;
  }

  const exportExcel = () => {
    if (tableRef.current === null) {
      alert("请先上传参保表格");
      return;
    }
    const wb = utils.book_new();
    const ws = utils.table_to_sheet(tableRef.current, {raw: true});

    utils.book_append_sheet(wb, ws, "Sheet1");
    writeFile(wb, fileName);
  }

  return (
    <div style={{display: "flex", padding: 8, paddingTop: 32, flexDirection: "column"}}>
      <span>1. 请上传参保表格:</span>
      <div style={{display: "flex", justifyContent: "space-between"}}>
        <input ref={excelInputRef} type={"file"} onChange={loadExcel}/>
        <button onClick={resetExcel}>清空参保表格</button>
      </div>

      <hr/>
      <span>2. 请上传要赛选的人员名单:</span>
      <div style={{display: "flex", justifyContent: "space-between"}}>
        <input ref={personInputRef} type={"file"} onChange={loadPerson}/>
        <button onClick={resetPerson}>清空人员名单</button>
      </div>


      <hr/>
      <button onClick={exportExcel}>3. 导出结果</button>
      {
        months.length != 0 &&
        <table border={1} ref={tableRef}>
            <thead>
            <tr>
                <th style={{width: 200}}>姓名</th>
                <th>身份证</th>
                <th>开始参保年月</th>
                <th>终止参保年月</th>
              {months.map((m, i) => <th key={`month_${i}`}>{yearMonthToString(m)}</th>)}
            </tr>
            </thead>
            <tbody>
            {
              Array.from(data.entries())
                .filter(([key]) => persons.length == 0 ? true : persons.some(p => Person.toStr(p) === key))
                .map(([key, value], i) => {
                  const [start, end] = sortYearMonth(value);
                  const person = Person.fromString(key);
                  return <tr key={`row_${i}`}>
                    <td>{person.name}</td>
                    <td>{person.id}</td>
                    <td>{yearMonthToDateString(start, true)}</td>
                    <td>{yearMonthToDateString(end, false)}</td>
                    {months.map((month, j) =>
                      <td key={`money_${i}_${j}`}
                          style={{textAlign: 'right'}}>{value.some(m => m === month) ? '650' : ''}</td>
                    )}
                  </tr>
                })
            }
            </tbody>
        </table>
      }
    </div>
  );
}

export default App;
