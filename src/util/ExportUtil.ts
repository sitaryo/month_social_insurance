import {utils, write, writeFile} from "xlsx";
import moment, {Moment} from "moment";
import JSZip from "jszip";

export class CompanyStatisticsRecord {
  name: string = '';
  idType: string = '';
  id: string = '';
  cardNum: string = '';
  cardType: string = '';
  socialInsuranceMonth: number = 0;
}

class ExportUtil {

  private static numberToLetters = (num: number) => {
    let letters = ''
    while (num >= 0) {
      letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[num % 26] + letters
      num = Math.floor(num / 26) - 1
    }
    return letters
  }

  private static collectToCompanyStatistics = (persons: Map<string, Set<string>>) => {
    const result: CompanyStatisticsRecord[] = [];
    for (let [nameAndId, months] of persons.entries()) {
      const [name, id] = nameAndId.split("_");
      result.push({
        cardNum: "", cardType: "", idType: "",
        name,
        id,
        socialInsuranceMonth: months.size
      });
    }
    return result;
  }

  private static getMonthRange = (start: Moment, end: Moment) => {
    const months = [start];
    while (months[months.length - 1].diff(end) < 0) {
      months.push(months[months.length - 1].clone().add(1, 'month'));
    }
    return months;
  }

  private static collectToPersonStatistics = (persons: Map<string, Set<string>>) => {
    const years =
      Array.from(persons.entries())
        .flatMap(([_, m]) => Array.from(m))
        .filter(m => !!m)
        .map(m => moment(m, 'YYYYMM', true))
        .filter(m => m.isValid())
        .sort((a, b) => a.diff(b));
    const start = years[0];
    const end = years[years.length - 1];

    const monthsRange = ExportUtil.getMonthRange(start, end)
    const totalRecord: any[] = [
      ['序号', '姓名', '身份证号', '证件编号', '证件类型', '开始参保年月', '结束参保年月', '参保总月数', '合计', ...monthsRange.map(m => m.format("yyyyMM"))],
    ];

    Array.from(persons.entries())
      .forEach(([p, m], i) => {
          const ms = Array.from(m).map(m => moment(m, 'YYYYMM'));
          const monthSorted = ms.sort((a, b) => a.diff(b));
          const s = monthSorted.length ? monthSorted[0] : null;
          const e = monthSorted.length ? monthSorted[monthSorted.length - 1] : null;
          const [name, id] = p.split("_");
          totalRecord.push([
            i + 1,
            name,
            id,
            '',
            '',
            s?.format('YYYY年MM月') || '',
            e?.format('YYYY年MM月') || '',
            {
              t: 'n',
              f: `counta(${ExportUtil.numberToLetters(9) + (i + 2)}:${ExportUtil.numberToLetters(9 + monthsRange.length - 1) + (i + 2)})`
            },
            {
              t: 'n',
              f: `sum(${ExportUtil.numberToLetters(9) + (i + 2)}:${ExportUtil.numberToLetters(9 + monthsRange.length - 1) + (i + 2)})`
            },
            ...monthsRange.map(m => ms.some(e => e.isSame(m)) ? 650 : null),
          ]);
        }
      )

    totalRecord.push([
      '合计',
      '',
      '',
      '',
      '',
      '',
      '',
      {
        t: 'n',
        f: `sum(${ExportUtil.numberToLetters(7) + 2}:${ExportUtil.numberToLetters(7) + totalRecord.length})`
      },
      {
        t: 'n',
        f: `sum(${ExportUtil.numberToLetters(8) + 2}:${ExportUtil.numberToLetters(8) + totalRecord.length})`
      },
      ...monthsRange.map((_, i) =>
        ({
          t: 'n',
          f: `sum(${ExportUtil.numberToLetters(9 + i) + 2}:${ExportUtil.numberToLetters(9 + i) + totalRecord.length})`
        })
      ),
    ]);

    return totalRecord;
  }

  static exportCompanyStatistics = (data: Map<string, Map<string, Set<string>>>) => {
    const wbs = new Map();
    Array.from(data.entries()).forEach(([company, persons]) => {
      const wb = utils.book_new();
      const totalRecord = [
        ['xh', 'lyrxm', 'sfzjlxDm', 'sfzjhm', 'jycyzbh', 'lxDm', 'zbqygzsj'],
        ['*序号', '*招用人姓名', '*身份证件类型', '*身份证件号码', '证件编号', '*类型(1)(2)(3)(4)', '*在本企业工作时间（月）'],
      ]

      ExportUtil
        .collectToCompanyStatistics(persons)
        .forEach((record, i) => {
          totalRecord.push([
            `${i + 1}`,
            record.name,
            record.idType,
            record.id,
            record.cardNum,
            record.cardType,
            record.socialInsuranceMonth + ""
          ]);
        });
      totalRecord.push(['结束标志']);

      const ws = utils.aoa_to_sheet(totalRecord);

      utils.book_append_sheet(wb, ws, "填入模板信息");

      const sheet2 = [
        ['序号', '身份证件类型代码', '身份证件种类名称', '上级身份证件类型代码'],
        ['1', '227', '中国护照', '200'],
        ['2', '228', '城镇退役士兵自谋职业证', '200'],
        ['3', '100', '单位', ''],
        ['4', '101', '组织机构代码证', '100'],
        ['5', '199', '其他证件', '100'],
        ['6', '200', '个人', ''],
        ['7', '201', '居民身份证', '200'],
        ['8', '202', '军官证', '200'],
        ['9', '203', '武警警官证', '200'],
        ['10', '204', '士兵证', '200'],
        ['11', '205', '军队离退休干部证', '200'],
        ['12', '206', '残疾人证', '200'],
        ['13', '207', '残疾军人证（1-8级）', '200'],
        ['14', '208', '外国护照', '200'],
        ['15', '209', '港澳同胞回乡证', '200'],
        ['16', '210', '港澳居民来往内地通行证', '200'],
        ['17', '211', '台胞证', '200'],
        ['18', '212', '中华人民共和国往来港澳通行证', '200'],
        ['19', '213', '台湾居民来往大陆通行证', '200'],
        ['20', '214', '大陆居民往来台湾通行证', '200'],
        ['21', '215', '外国人居留证', '200'],
        ['22', '216', '外交官证', '200'],
        ['23', '217', '领事馆证', '200'],
        ['24', '218', '海员证', '200'],
        ['25', '219', '香港身份证', '200'],
        ['26', '220', '台湾身份证', '200'],
        ['27', '221', '澳门身份证', '200'],
        ['28', '222', '外国人身份证件', '200'],
        ['29', '223', '高校毕业生自主创业证', '200'],
        ['30', '224', '就业失业登记证', '200'],
        ['31', '225', '退休证', '200'],
        ['32', '226', '离休证', '200'],
        ['33', '299', '其他个人证件', '200'],
      ];

      const ws2 = utils.aoa_to_sheet(sheet2);

      utils.book_append_sheet(wb, ws2, "身份证件类型代码");

      const sheet3 = [
        ['序号', '类型代码', '类型名称'],
        ['1', '01', '在人力资源社会保障部门公共就业服务机构登记失业半年以上人员'],
        ['2', '02', '零就业家庭、享受城市居民最低生活保障家庭劳动年龄内的登记失业人员'],
        ['3', '03', '毕业年度内高校毕业生'],
        ['4', '04', '纳入全国扶贫开发信息系统的农村建档立卡贫困人员'],
      ];

      const ws3 = utils.aoa_to_sheet(sheet3);

      utils.book_append_sheet(wb, ws3, "类型代码");

      const ws4 = utils.aoa_to_sheet(ExportUtil.collectToPersonStatistics(persons));
      utils.book_append_sheet(wb, ws4, "个人统计");

      wbs.set(`${company}[汇总].xlsx`, wb);
    })
    const zip = new JSZip();
    Array.from(wbs.entries())
      .forEach(([name, wb]) => zip.file(name, write(wb, {type: "buffer"})));
    const download = (file: BlobPart, name?: string) => {
      const url = URL.createObjectURL(new Blob([file]));
      const dl = document.createElement('a');
      dl.download = name || ('fflate-demo-' + Date.now() + '.dat');
      dl.href = url;
      dl.click();
      URL.revokeObjectURL(url);
    }
    console.log(wbs.size);
    zip.generateAsync({type: "blob"})
      .then(file => {
        download(file, "汇总.zip");
      });
  }

  static exportPersonStatistics = (company: string, data: Map<string, Set<string>>) => {
    const wb = utils.book_new();
    const ws = utils.aoa_to_sheet(ExportUtil.collectToPersonStatistics(data));

    utils.book_append_sheet(wb, ws, "个人统计");

    writeFile(wb, `${company}[汇总].xlsx`, {});
  }
}

export default ExportUtil;
