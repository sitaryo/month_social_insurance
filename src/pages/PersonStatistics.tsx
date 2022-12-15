import React, {ChangeEvent, useRef, useState} from 'react';
import {read, utils} from "xlsx";
import ExportUtil from "../util/ExportUtil";

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


interface P {

}

const PersonStatistics: React.FC<P> = (props) => {
  const [records, setRecords] = useState<Map<string, Set<string>>>(new Map());
  const [persons, setPersons] = useState<Person[]>([]);
  const [fileName, setFileName] = useState("");
  const [companyName, setCompanyName] = useState("");
  const [excelError, setExcelError] = useState('');
  const [personError, setPersonError] = useState('');
  const [loading, setLoading] = useState(false);
  const [exportType, setExportType] = useState<ExportType>("excel");

  const personInputRef = useRef<any>(null);
  const excelInputRef = useRef<any>(null);

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
        const data = new Map<string, Set<string>>();
        persons.forEach(p => data.set(Person.toStr(p), new Set()));

        const res = f.target?.result as string;
        const workBook = read(res, {type: "binary"});
        workBook.SheetNames.forEach(sheetName => {
          const month = sheetName.trim();
          const ds = utils.sheet_to_json<Person>(workBook.Sheets[sheetName], {header: ['name', 'id']});
          ds.splice(0, 1);
          ds.forEach(row => {
            try {
              row.name = `${row?.name ?? ""}`.trim();
              row.id = `${row?.id ?? ""}`.trim();
              const key = Person.toStr(row);

              // 同一个人可能在同一个sheet内存在多条数据，every 去重
              if (data.has(key)) {
                data.get(key)?.add(month);
              }
            } catch (e) {
              setLoading(false);
              setExcelError(`读取参保数据错误：sheet[${sheetName}] 中格式错误,内容【 姓名${row.name} 身份证${row.id}】`);
              throw e;
            }
          });
        })

        setRecords(data);
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
    setRecords(new Map());
    setFileName("");
    excelInputRef.current.value = null;
  }

  const exportExcel = () => ExportUtil.exportPersonStatistics(companyName, collectRecord(records, persons))

  const changeCompanyName = (event: ChangeEvent<HTMLInputElement>) => setCompanyName(event.target.value ?? "")

  const changeExportType = (e: ChangeEvent<HTMLSelectElement>) => setExportType(e.target.value as ExportType)

  const collectRecord = (records: Map<string, Set<string>>, persons: Person[]): Map<string, Set<string>> => {
    const result: Map<string, Set<string>> = new Map();
    const personsToFilter = new Set(persons.map(p => Person.toStr(p)));
    if (exportType === "person") {
      Array.from(personsToFilter).forEach(p => {
        result.set(p, records.get(p) || new Set());
      });
    } else {
      for (let [p, m] of records.entries()) {
        if (personsToFilter.has(p)) {
          result.set(p, m);
        }
      }
    }

    return result;
  }

  return (
    <React.Fragment>
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
        <option value={"excel"}>以参保表为准</option>
        <option value={"person"}>以人员表为准，人员表有几个人汇总表就有几项</option>
      </select>

      <hr/>
      <button onClick={exportExcel}>5. 导出结果</button>
      {excelError && <h5>{excelError}</h5>}
      {personError && <h5>{personError}</h5>}
      {loading && <h3 style={{textAlign: "center", color: 'red'}}>正在处理数据</h3>}
    </React.Fragment>
  );
}

export default PersonStatistics;
