import React, {ChangeEvent, useMemo, useRef, useState} from 'react';
import {read, utils, writeFile} from "xlsx";
import moment, {Moment} from "moment";

type ExportType = "person" | "excel";

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
  const [companyName, setCompanyName] = useState("");
  const [excelError, setExcelError] = useState('');
  const [personError, setPersonError] = useState('');
  const [loading, setLoading] = useState(false);
  const [exportType, setExportType] = useState<ExportType>("person");

  const personInputRef = useRef<any>(null);
  const excelInputRef = useRef<any>(null);
  const tableRef = useRef<any>(null);


  const readStringAsYearMonth = (s: string) => moment(s, 'YYYYMM');

  const loadExcel = (event: ChangeEvent<HTMLInputElement>) => {
    if (persons.length === 0) {
      alert("请先上传人员名单");
      resetExcel();
      return;
    }
    const files = event.target.files;
    setExcelError("");
    if (files) {
      setFileName("[汇总]" + files[0].name);
      let fr = new FileReader();
      fr.readAsBinaryString(files[0]);
      setLoading(true);
      fr.onload = (f) => {
        const data = new Map<string, Moment[]>();
        persons.forEach(p => data.set(Person.toStr(p), []));
        const months: Moment[] = [];

        const res = f.target?.result as string;
        const workBook = read(res, {type: "binary"});
        workBook.SheetNames.forEach(sheetName => {
          const ds = utils.sheet_to_json<Person>(workBook.Sheets[sheetName], {header: ['name', 'id']});
          const month = readStringAsYearMonth(sheetName);
          months.push(month);
          ds.splice(0, 1);
          ds.forEach(row => {
            try {
              row.name = `${row?.name ?? ""}`.trim();
              row.id = `${row?.id ?? ""}`.trim();
              const key = Person.toStr(row);

              // 同一个人可能在同一个sheet内存在多条数据，every 去重
              if (data.has(key) && data.get(key)?.every(m => m.diff(month) != 0)) {
                data.get(key)?.push(month);
              }
            } catch (e) {
              setLoading(false);
              setExcelError(`读取参保数据错误：sheet[${sheetName}] 中格式错误,内容【 姓名${row.name} 身份证${row.id}】`);
              throw e;
            }
          });
        })

        setMonths(months.sort((a, b) => a.diff(b)));
        setData(data);
        setLoading(false);
      };
    }
  }

  const loadPerson = (event: ChangeEvent<HTMLInputElement>) => {
    setPersonError("");
    const files = event.target.files;
    if (files) {
      setLoading(true);
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
          try {
            p.name = `${p?.name ?? ""}`.trim();
            p.id = `${p?.id ?? ""}`.trim();
          } catch (e) {
            setLoading(false);
            setPersonError(`读取参保数据错误：内容【 姓名${p.name} 身份证${p.id}】`);
            throw e;
          }
        });
        setPersons(ds);
        setLoading(false);
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
    const wb = utils.book_new();
    const ws = utils.table_to_sheet(tableRef.current, {raw: true});

    utils.book_append_sheet(wb, ws, "Sheet1");
    writeFile(wb, fileName);
  }

  const changeCompanyName = (event: ChangeEvent<HTMLInputElement>) => {
    setCompanyName(event.target.value ?? "");
  }

  const changeExportType = (e: ChangeEvent<HTMLSelectElement>) => {
    setLoading(true);
    setExportType(e.target.value as ExportType);
    setTimeout(() => {
      setLoading(false);
    }, 500);
  }

  return (
    <div style={{display: "flex", padding: 8, paddingTop: 32, flexDirection: "column"}}>
      <h3 style={{textAlign: "center"}}>参保汇总工具 v1.0.6</h3>

      <span>1. 请上传要筛选的人员名单:</span>
      <div style={{display: "flex", justifyContent: "space-between"}}>
        <input ref={personInputRef} type={"file"} onChange={loadPerson}/>
        <button onClick={resetPerson}>清空人员名单</button>
      </div>

      <hr/>
      <span>2. 请上传参保表格:</span>
      <div style={{display: "flex", justifyContent: "space-between"}}>
        <input ref={excelInputRef} type={"file"} onChange={loadExcel}/>
        <button onClick={resetExcel}>清空参保表格</button>
      </div>

      <hr/>
      <span>3. 公司名称:</span>
      <input type={"text"} value={companyName} onChange={changeCompanyName}/>

      <hr/>
      <span>4. 导出方式:</span>
      <select value={exportType} onChange={changeExportType}>
        <option value={"person"}>以人员表为准，人员表有几个人汇总表就有几项</option>
        <option value={"excel"}>以参保表为准</option>
      </select>

      <hr/>
      <button onClick={exportExcel}>5. 导出结果</button>
      {excelError && <h5>{excelError}</h5>}
      {personError && <h5>{personError}</h5>}
      {
        loading ?
          <h3 style={{textAlign: "center", color: 'red'}}>正在处理数据</h3> :
          <Table months={months} data={data} persons={persons} companyName={companyName} table={tableRef}
                 exportType={exportType}/>
      }
    </div>
  );
}

interface P {
  months: Moment[],
  data: Map<string, Moment[]>,
  persons: Person[],
  companyName: string,
  table: any,
  exportType: ExportType,
}

const Table: React.FC<P> = (props) => {
  const {months, data, persons, companyName, table, exportType} = props;

  const years = () => Array.from(new Set(months.map(m => m.year()))).sort((a, b) => a - b);

  const renderYearColumn = () => {
    const year = new Map<number, number>();
    months.forEach(month => {
      if (year.has(month.year())) {
        year.set(month.year(), (year.get(month.year()) ?? 0) + 1);
      } else {
        year.set(month.year(), 1);
      }
    });
    return Array.from(year.entries())
      .sort((a, b) => a[0] - b[0])
      .map(([year, count], i) =>
        <th key={`year_col_${i}`} colSpan={count}>{year}</th>
      );
  }

  const footer = () => {
    const personToCount = Array.from(data.entries())
      .filter(([key]) => persons.length == 0 ? true : persons.some(p => Person.toStr(p) === key))
      .map(([_, v]) => v)
      .flat();

    return months.map((m, i) => <td key={`month_${i}`}>{personToCount.filter(s => s.diff(m) == 0).length * 650}</td>)
  }

  const monthCount = () => {
    const personToCount = Array.from(data.entries())
      .filter(([key]) => persons.length == 0 ? true : persons.some(p => Person.toStr(p) === key))
      .map(([_, v]) => v)
      .flat();
    return personToCount.length;
  }

  const yearCount = () => {
    const personToCount = Array.from(data.entries())
      .filter(([key]) => persons.length == 0 ? true : persons.some(p => Person.toStr(p) === key))
      .map(([_, v]) => v)
      .flat();
    return years().map((y, i) => <td
      key={`year_count_${i}`}>{personToCount.filter(v => v.year() === y).length * 650}</td>);
  }

  const yearMonthToString = (yearMonth: Moment | undefined) => yearMonth?.format('YYYYMM') ?? "";

  const sortYearMonth = (yearMonths: Moment[]) => {
    const sort = yearMonths.sort((a, b) => a.diff(b));
    return [sort[0], sort[sort.length - 1]]
  }

  const renderRow = (person: Person, value: Moment[], i: number) => {
    const [start, end] = sortYearMonth(value);
    return <tr key={`row_${i}`}>
      <td>{person.name}</td>
      <td>{person.id}</td>
      <td>{yearMonthToString(start)}</td>
      <td>{yearMonthToString(end)}</td>
      {months.map((month, j) =>
        <td key={`money_${i}_${j}`}
            style={{textAlign: 'right'}}>{value.some(m => m === month) ? '650' : ''}</td>
      )}
      <td>{value.length}</td>
      {
        years().map((y, j) => <td
          key={`year_count_${i}_${j}`}>{value.filter(v => v.year() == y).length * 650}</td>)
      }
      <td>{value.length * 650}</td>
    </tr>
  }

  return <table border={1} ref={table}>
    <thead>
    <tr>
      <th>{companyName}</th>
      <th/>
      <th/>
      <th/>
      {useMemo(() => renderYearColumn(), [data, months, persons])}
      <th colSpan={2 + years().length}>数据统计</th>
    </tr>
    {useMemo(() =>
      <tr>
        <th style={{width: 200}}>姓名</th>
        <th>身份证</th>
        <th>开始参保年月</th>
        <th>终止参保年月</th>
        {months.map((m, i) => <th key={`month_${i}`}>{yearMonthToString(m)}</th>)}
        <th>购买社保合计月</th>
        {years().map(y => <th key={`year_${y}`}>{y + "年"}</th>)}
        <th>合计</th>
      </tr>, [data, months, persons])}
    </thead>
    {useMemo(() => <>
      <tbody>
      {
        exportType === "person" ?
          persons.map((person, i) => renderRow(person, data.get(Person.toStr(person)) ?? [], i)) :
          Array.from(data.entries())
            .filter(([key, value]) => persons.some(p => Person.toStr(p) === key) && value.length !== 0)
            .map(([key, value], i) => renderRow(Person.fromString(key), value, i))
      }
      </tbody>
      <tfoot>
      <tr>
        <td/>
        <td/>
        <td/>
        <td>合计</td>
        {footer()}
        <td>{monthCount()}</td>
        {yearCount()}
        <td>{monthCount() * 650}</td>
      </tr>
      </tfoot>
    </>, [data, months, persons, exportType])}
  </table>;
}


export default App;
