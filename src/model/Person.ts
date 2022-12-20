export const PersonToStr = (p: Person) => `${p.name}_${p.id}`;
export const strToPerson = (s:string)=>s.split("_");
export interface Person {
  name: string;
  id: string;
}
