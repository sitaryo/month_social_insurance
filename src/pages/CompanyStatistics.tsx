import React, {ChangeEvent, useRef, useState} from 'react';
import {read, utils} from "xlsx";
import ExportUtil from "../util/ExportUtil";
import moment from "moment";
import {Person} from "./PersonStatistics";

class Record {
  company: string = '';
  month: string = '';
  name: string = '';
  id: string = '';
}

interface P {
}

const recordToPersonString = (record: Record) => `${record.name}_${record.id}`;
const personStringToPerson = (str: string) => str.split("_");

const CompanyStatistics: React.FC<P> = (props) => {
  const [records, setRecords] = useState<Record[]>([]);
  const [company, setCompany] = useState<Set<string>>(new Set());
  const [persons, setPersons] = useState<Set<string>>(new Set());
  const [loading, setLoading] = useState(false);
  const [errors, setErrors] = useState<string[]>([]);

  const excelRef = useRef<any>();
  const personsRef = useRef<any>();
  const companyRef = useRef<any>();

  const resetExcel = () => {
    setRecords([]);
    setErrors([]);
    excelRef.current.value = null;
  }
  const resetCompany = () => {
    setCompany(new Set());
    setErrors([]);
    companyRef.current.value = null;
  }


  const uploadExcel = (event: ChangeEvent<HTMLInputElement>) => {
    if (!company.size) {
      alert("请先上传公司表");
      return;
    }
    if (!persons.size && !window.confirm("未上传人员表，是否继续？")) {
      console.log('return')
      resetExcel();
      return;
    }
    console.log('true')
    const files = event.target.files;
    if (files) {
      let total = files.length;
      let totalRecord: Record[] = [];
      const errMsg: string[] = [];
      let process = 0;
      setLoading(true);

      for (let i = 0; i < files.length; i++) {
        let fr = new FileReader();
        const filename = files[i].name;
        fr.readAsBinaryString(files[i]);
        fr.onload = (f) => {
          const res = f.target?.result as string;
          const workBook = read(res, {type: "binary"});
          workBook.SheetNames.forEach((sheetName, i) => {
            const ds = utils.sheet_to_json<Record>(
              workBook.Sheets[sheetName],
              {header: ['company', 'month', 'name', 'id'], blankrows: false}
            );
            ds.splice(0, 1);
            const result = ds.map((p, i) => {
              p.company = `${p?.company ?? ""}`.trim();
              p.month = `${p?.month ?? ""}`.trim();
              if (!moment(p.month, 'YYYYMM', true).isValid()) {
                errMsg.push(filename + " 文件 : sheet :" + sheetName + " " + (i + 2) + " 行时间格式错误： " + p.month);
              }
              p.name = `${p?.name ?? ""}`.trim().replaceAll(" ", "");
              p.id = `${p?.id ?? ""}`.trim();
              return p;
            }).filter(r => (persons.size == 0 || persons.has(recordToPersonString(r))) && company.has(r.company));
            totalRecord = [...totalRecord, ...result];
          });

          process++;
          if (process === total) {
            setLoading(false);
            setRecords(totalRecord);
            setErrors(errMsg);
          }
        };
      }
    }
  }

  const uploadCompany = (event: ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (files) {
      setLoading(true);
      let fr = new FileReader();
      fr.readAsBinaryString(files[0]);
      fr.onload = (f) => {
        const res = f.target?.result as string;
        const workBook = read(res, {type: "binary"});
        const ds = utils.sheet_to_json<any>(
          workBook.Sheets[workBook.SheetNames[0]],
          {header: ['name']}
        );
        ds.splice(0, 1);
        const result = ds.map(c => `${c.name ?? ""}`.trim());
        setCompany(new Set(result));
        setLoading(false);
      };
    }
  }

  const collectRecord = (records: Record[], company: Set<string>): Map<string, Map<string, Set<string>>> => {
    const companyMap = new Map<string, Map<string, Set<string>>>();
    records.forEach(r => {
      // 排除非需要导出的公司
      if (!company.has(r.company)) {
        return;
      }
      const personStr = recordToPersonString(r);
      if (companyMap.has(r.company)) {
        const persons = companyMap.get(r.company)!;
        if (persons.has(personStr)) {
          persons.get(personStr)!.add(r.month);
        } else {
          persons.set(personStr, new Set([r.month]));
        }
      } else {
        const some = new Map<string, Set<string>>();
        some.set(personStr, new Set([r.month]));
        companyMap.set(r.company, some);
      }
    });
    return companyMap;
  }

  const exportToExcel = () => {
    if (errors.length) {
      alert("请确保表格内容格式正确无误");
      return;
    }
    if (!records.length) {
      alert("请先上传参保表");
      return;
    }
    if (!company.size) {
      alert("请先上传公司表");
      return;
    }

    setLoading(true);
    const companyMap = collectRecord(records, company);
    ExportUtil.exportCompanyStatistics(companyMap);

    setLoading(false);
  }

  const loadPerson = (event: ChangeEvent<HTMLInputElement>) => {
    resetExcel();
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
        const result = ds.map(p => {
          p.name = `${p?.name ?? ""}`.trim().replaceAll(" ", "");
          p.id = `${p?.id ?? ""}`.trim();
          return Person.toStr(p);
        });
        setPersons(new Set(result));
        setLoading(false);
      };
    }
  }

  const resetPerson = () => {
    setPersons(new Set());
    setErrors([]);
    personsRef.current.value = null;
  }

  return <React.Fragment>

    <span>1. 请上传公司表: {company.size !== 0 && "已导入" + company.size + "条公司数据"}</span>
    <div style={{display: "flex", justifyContent: "space-between"}}>
      <input ref={companyRef} type={"file"} onChange={uploadCompany}/>
      <button onClick={resetCompany}>清空公司表</button>
    </div>

    <span>2. 请上传人员表: {persons.size !== 0 && "已导入" + persons.size + "条人员数据"}</span>
    <div style={{display: "flex", justifyContent: "space-between"}}>
      <input ref={personsRef} type={"file"} onChange={loadPerson}/>
      <button onClick={resetPerson}>清空人员表</button>
    </div>

    <span>3. 请上传参保表: {records.length !== 0 && "筛选出" + records.length + "条数据"}</span>
    <div style={{display: "flex", justifyContent: "space-between"}}>
      <input ref={excelRef} type={"file"} onChange={uploadExcel} multiple/>
      <button onClick={resetExcel}>清空参保表</button>
    </div>

    <hr/>
    <button onClick={exportToExcel}>3. 导出结果</button>

    {loading && <h3 style={{textAlign: "center", color: 'red'}}>正在处理数据</h3>}
    {
      !!errors.length &&
      <div style={{textAlign: "center", color: 'red'}}>发现{errors.length}条错误，以下为前1000条：</div>
    }
    {
      errors
        .slice(0, 1000)
        .map((e, i) => <div key={`error_${i}`} style={{textAlign: "center", color: 'red'}}>{e}</div>)
    }
  </React.Fragment>;
}

export default CompanyStatistics;
