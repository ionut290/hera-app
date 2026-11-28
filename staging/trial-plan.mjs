// V2 staging only. Server rules independently validate every quota.
export const TRIAL_DAYS=30;
export const TRIAL_LIMITS=Object.freeze({commesse:5,impianti:102,admin:1,operatori:11,mezzi:11,squadre:2,storageBytes:100000000,personale:11});
export function initialTrialUsage(){return {commesse:0,impianti:0,admin:1,operatori:0,mezzi:0,squadre:0,storageBytes:0,personale:0};}
export function quotaKind(path){
 const p=String(path).split('/');
 if(p.length===2){
  if(['commesse','neve_commesse','globalCommesse'].includes(p[0]))return 'commesse';
  if(['mezzi','neve_mezzi'].includes(p[0]))return 'mezzi';
  if(['personale','neve_personale'].includes(p[0]))return 'personale';
  if(['squadreCommesse','neve_squadre'].includes(p[0]))return 'squadre';
 }
 if(p.length===4&&['commesse','neve_commesse','globalCommesse'].includes(p[0])&&p[2]==='impianti')return 'impianti';
 return null;
}
