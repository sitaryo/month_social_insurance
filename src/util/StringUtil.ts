export default class StringUtil {
  static trim = (s: any | undefined) => `${s ?? ""}`.trim();

  static replaceBlank = (s: any | undefined) =>
    StringUtil.trim(s).replace(/ /g, '');
}
