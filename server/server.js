import "dotenv/config";
import express from "express";
import cors from "cors";
import { ClientSecretCredential } from "@azure/identity";
import { createTeamsNotificationService } from "./teams-notifications.js";

const app=express();
const PORT=Number(process.env.PORT||3001);
const GRAPH=process.env.GRAPH_BASE_URL||"https://graph.microsoft.com/v1.0";
const GROUP_NAME=process.env.PRESENCE_GROUP_NAME||"ShowInPresenceBoard";
const MEETING_ROOM_COMPANY="__meeting_rooms__";
const MISSING_COMPANY="__missing_company__";
const MISSING_DEPARTMENT="__missing_department__";
const credential=new ClientSecretCredential(process.env.TENANT_ID,process.env.CLIENT_ID,process.env.CLIENT_SECRET);
app.use(cors()); app.use(express.json());

let memberCache={expires:0,group:null,members:[]};
async function token(){const t=await credential.getToken("https://graph.microsoft.com/.default");if(!t?.token)throw new Error("Kein Graph-Token erhalten.");return t.token}
async function graph(path,options={}){
  const response=await fetch(path.startsWith("http")?path:`${GRAPH}${path}`,{...options,headers:{Authorization:`Bearer ${await token()}`,Accept:"application/json",...(options.body?{"Content-Type":"application/json"}:{}),...options.headers}});
  if(!response.ok){const text=await response.text();throw new Error(`Graph ${response.status}: ${text}`)}
  if(response.status===204)return null; return response.json();
}
async function collection(path,headers={}){const all=[];let next=path;while(next){const d=await graph(next,{headers});all.push(...(d.value||[]));next=d["@odata.nextLink"]||null}return all}
async function getMembers(){
  if(Date.now()<memberCache.expires)return memberCache;
  const escaped=GROUP_NAME.replaceAll("'","''");
  const groups=await graph(`/groups?$filter=${encodeURIComponent(`displayName eq '${escaped}'`)}&$select=id,displayName`);
  if((groups.value||[]).length!==1)throw new Error(`Gruppe ${GROUP_NAME} wurde nicht eindeutig gefunden.`);
  const group=groups.value[0];
  const members=(await collection(`/groups/${group.id}/transitiveMembers/microsoft.graph.user?$select=id,displayName,givenName,surname,mail,userPrincipalName,jobTitle,department,companyName,officeLocation,accountEnabled`)).filter(u=>u.accountEnabled!==false);
  memberCache={expires:Date.now()+86400000,group,members};return memberCache;
}
function chunks(a,n){const out=[];for(let i=0;i<a.length;i+=n)out.push(a.slice(i,i+n));return out}
async function presences(ids){const out=[];for(const batch of chunks(ids,650)){const d=await graph("/communications/getPresencesByUserId",{method:"POST",body:JSON.stringify({ids:batch})});out.push(...(d.value||[]))}return out}
function formatTime(value){if(!value)return null;const d=new Date(value);return Number.isNaN(d.getTime())?null:new Intl.DateTimeFormat("de-DE",{hour:"2-digit",minute:"2-digit",hour12:false,timeZone:"Europe/Berlin"}).format(d)}
function formatDate(value){if(!value)return null;const d=new Date(value);return Number.isNaN(d.getTime())?null:new Intl.DateTimeFormat("de-DE",{day:"2-digit",month:"2-digit",year:"numeric",timeZone:"Europe/Berlin"}).format(d)}
function formatDateTime(value){if(!value)return null;const d=new Date(value);return Number.isNaN(d.getTime())?null:new Intl.DateTimeFormat("de-DE",{day:"2-digit",month:"2-digit",year:"numeric",hour:"2-digit",minute:"2-digit",timeZone:"Europe/Berlin"}).format(d)}
function parseGraphDateTime(value){
  const dateTime=typeof value==="string"?value:value?.dateTime;
  if(!dateTime)return new Date(Number.NaN);
  if(/(?:Z|[+-]\d{2}:\d{2})$/i.test(dateTime))return new Date(dateTime);
  const wallTime=new Date(`${dateTime}Z`);
  if(Number.isNaN(wallTime.getTime()))return wallTime;
  const rawZone=typeof value==="object"?value.timeZone||"UTC":"UTC";
  const zoneAliases={"w. europe standard time":"Europe/Berlin","romance standard time":"Europe/Paris"};
  const timeZone=zoneAliases[rawZone.toLowerCase()]||rawZone;
  if(["utc","gmt","etc/utc"].includes(timeZone.toLowerCase()))return wallTime;
  try{
    const formatter=new Intl.DateTimeFormat("en-US",{timeZone,hourCycle:"h23",year:"numeric",month:"2-digit",day:"2-digit",hour:"2-digit",minute:"2-digit",second:"2-digit"});
    let timestamp=wallTime.getTime();
    for(let attempt=0;attempt<2;attempt++){
      const parts=Object.fromEntries(formatter.formatToParts(new Date(timestamp)).map(part=>[part.type,part.value]));
      const representedAsUtc=Date.UTC(Number(parts.year),Number(parts.month)-1,Number(parts.day),Number(parts.hour),Number(parts.minute),Number(parts.second));
      timestamp=wallTime.getTime()-(representedAsUtc-timestamp);
    }
    return new Date(timestamp);
  }catch{
    return new Date(dateTime);
  }
}
async function calendar(userId,rangeDays=1){
  const now=new Date(),start=new Date(now),end=new Date(now);
  start.setDate(start.getDate()-1);
  end.setDate(end.getDate()+rangeDays);
  const path=`/users/${userId}/calendar/calendarView?startDateTime=${encodeURIComponent(start.toISOString())}&endDateTime=${encodeURIComponent(end.toISOString())}&$select=id,start,end,showAs,isAllDay,isCancelled&$orderby=start/dateTime`;
  return collection(path,{Prefer:'outlook.timezone="Europe/Berlin"'});
}
async function mailboxSettings(userId){
  return graph(`/users/${encodeURIComponent(userId)}/mailboxSettings/automaticRepliesSetting`,{headers:{Prefer:'outlook.timezone="UTC"'}});
}
async function roomPlaces(){
  return collection("/places/microsoft.graph.room?$select=emailAddress,capacity,building,floorNumber,floorLabel,label");
}
function mapCalendar(events){
  const now=new Date(); const list=(events||[]).filter(e=>!e.isCancelled).map(e=>({...e,s:parseGraphDateTime(e.start),e:parseGraphDateTime(e.end)})).filter(e=>!Number.isNaN(e.s.getTime())&&!Number.isNaN(e.e.getTime())).sort((a,b)=>a.s-b.s);
  const today=formatDate(now);
  const oof=list.find(e=>e.showAs==="oof"&&e.s<=now&&e.e>now);
  const current=list.find(e=>e.showAs!=="oof"&&!e.isAllDay&&e.s<=now&&e.e>now);
  const next=list.find(e=>e.showAs!=="oof"&&!e.isAllDay&&e.s>now&&formatDate(e.s)===today);
  const schedule=list.map(event=>({start:event.s.toISOString(),end:event.e.toISOString(),showAs:event.showAs||"busy",isAllDay:Boolean(event.isAllDay)}));
  return {currentMeetingEnd:current?formatTime(current.e):null,currentMeetingShowAs:current?.showAs||null,meetingReachable:current?current.showAs==="free":true,nextMeeting:!current&&next?formatTime(next.s):null,oofAbsence:oof?{returnDate:formatDate(oof.e)}:null,schedule};
}
function plainText(value=""){
  return String(value)
    .replace(/<br\s*\/?\s*>/gi," ")
    .replace(/<\/p\s*>/gi," ")
    .replace(/<[^>]*>/g," ")
    .replace(/&nbsp;|&#160;/gi," ")
    .replace(/&amp;/gi,"&")
    .replace(/&lt;/gi,"<")
    .replace(/&gt;/gi,">")
    .replace(/&quot;/gi,'"')
    .replace(/&#39;|&apos;/gi,"'")
    .replace(/\s+/g," ")
    .trim();
}
function mapAutomaticReplies(settings){
  if(!settings)return {configured:false,status:"unavailable",active:false,note:""};
  const replies=settings.automaticRepliesSetting||settings;
  if(!replies)return {configured:false,status:"missing",active:false,note:""};
  const status=(replies?.status||"").trim().toLowerCase();
  const startDateTime=replies.scheduledStartDateTime;
  const endDateTime=replies.scheduledEndDateTime;
  const internalNote=plainText(replies.internalReplyMessage||"");
  const externalNote=plainText(replies.externalReplyMessage||replies.externalRelayMessage||"");
  const note=internalNote||externalNote;
  if(status==="disabled")return {configured:true,status,active:false,note};
  if(status==="alwaysenabled"||status==="enabled")return {configured:true,status,active:true,returnDate:null,note};
  if(status!=="scheduled")return {configured:false,status:status||"unknown",active:false,note};
  const hasStart=Boolean(startDateTime?.dateTime);
  const hasEnd=Boolean(endDateTime?.dateTime);
  const start=parseGraphDateTime(startDateTime);
  const end=parseGraphDateTime(endDateTime);
  const now=new Date();
  if(!hasStart&&!hasEnd)return {configured:true,status,active:true,returnDate:null,note};
  if((hasStart&&Number.isNaN(start.getTime()))||(hasEnd&&Number.isNaN(end.getTime())))return {configured:true,status,active:false,note};
  if((hasStart&&start>now)||(hasEnd&&end<=now))return {configured:true,status,active:false,note};
  return {configured:true,status,active:true,returnDate:hasEnd?formatDateTime(end):null,note};
}
function mapPresence(p){
  const availability=p?.availability||"PresenceUnknown",activity=p?.activity||"Offline";
  const pm={Available:"available",AvailableIdle:"available",Busy:"busy",BusyIdle:"busy",DoNotDisturb:"dnd",Away:"away",BeRightBack:"away",Offline:"offline",PresenceUnknown:"offline"};
  const outOfOffice=[availability,activity].some(value=>value.toLowerCase().replace(/[\s_-]/g,"")==="outofoffice");
  let presence=pm[availability]||"offline";if(!outOfOffice&&["InAMeeting","Presenting"].includes(activity))presence="meeting";
  const phone=["InACall","InAConferenceCall"].includes(activity)?"call":"free";
  const rawLocationType=
    p?.workLocation?.workLocationType||
    p?.workLocation?.type||
    p?.workLocationType||
    p?.location||
    "";
  const type=typeof rawLocationType==="string"?rawLocationType.toLowerCase().replace(/[\s_-]/g,""):"";
  const location=["remote","home","homeoffice","workfromhome","wfh"].includes(type)
    ?"home"
    :["office","onsite","onpremises"].includes(type)
      ?"office"
      :null;
  return {presence,phone,availability,activity,location,outOfOffice};
}
function initials(u){if(u.givenName||u.surname)return `${u.givenName?.[0]||""}${u.surname?.[0]||""}`.toUpperCase();return (u.displayName||"").replace(","," ").split(/\s+/).filter(Boolean).slice(0,2).map(x=>x[0]).join("").toUpperCase()}
function company(name=""){const s=String(name||"").trim().toLowerCase();if(s.includes("meeting room"))return MEETING_ROOM_COMPANY;if(!s)return MISSING_COMPANY;if(s.includes("global aviation")||s.includes("gapp"))return"GAPP";if(s.includes("european aviation")||s.includes("eac2")||s.includes("eacc"))return"EAC2";return"PAG"}
async function build(){
  const {group,members}=await getMembers(); const ps=await presences(members.map(x=>x.id)); const byId=new Map(ps.map(x=>[x.id,x]));
  const roomMembers=members.filter(u=>company(u.companyName)===MEETING_ROOM_COMPANY);
  const calendarResults=await Promise.allSettled(members.map(async u=>[u.id,await calendar(u.id,company(u.companyName)===MEETING_ROOM_COMPANY?30:1)]));const calendars=new Map(calendarResults.filter(x=>x.status==="fulfilled").map(x=>x.value));
  const roomIds=new Set(roomMembers.map(u=>u.id));
  const failedRoomCalendars=calendarResults.filter((result,index)=>result.status==="rejected"&&roomIds.has(members[index].id));
  if(failedRoomCalendars.length)console.warn(`[presence-board] Calendar.Read failed for ${failedRoomCalendars.length}/${roomMembers.length} rooms: ${failedRoomCalendars[0].reason?.message||"unknown Graph error"}`);
  const places=new Map();
  if(roomMembers.length){
    try{
      const results=await roomPlaces();
      for(const place of results){if(place.emailAddress)places.set(place.emailAddress.toLowerCase(),place)}
    }catch(error){
      console.warn(`[presence-board] Places.Read failed: ${error.message||"unknown Graph error"}`);
    }
  }
  const mailboxResults=await Promise.allSettled(members.map(async u=>[u.id,await mailboxSettings(u.id)]));
  const mailboxFailures=mailboxResults.flatMap((result,index)=>result.status==="rejected"?[{user:members[index],error:result.reason}]:[]);
  if(mailboxFailures.length){
    const details=mailboxFailures.map(({user,error})=>{
      const mailbox=user.mail||user.userPrincipalName;
      const identity=user.displayName||mailbox||user.id;
      return `${identity}${mailbox&&mailbox!==identity?` (${mailbox})`:""}: ${error?.message||"unknown Graph error"}`;
    }).join("; ");
    console.warn(`[presence-board] MailboxSettings.Read failed for ${mailboxFailures.length}/${members.length} users: ${details}`);
  }
  const mailboxes=new Map(mailboxResults.filter(x=>x.status==="fulfilled").map(x=>x.value));
  const people=members.map(u=>{
    const p=mapPresence(byId.get(u.id)),c=mapCalendar(calendars.get(u.id)||[]),sageAbsence=null,companyKey=company(u.companyName),place=places.get((u.mail||u.userPrincipalName||"").toLowerCase());
    const resourceLocationDescription=[u.officeLocation,place?.building,place?.floorLabel||place?.floorNumber,place?.label].filter(Boolean).map(String).filter((value,index,values)=>values.indexOf(value)===index).join(" · ");
    const automaticReplies=mapAutomaticReplies(mailboxes.get(u.id));
    const automaticRepliesAbsence=automaticReplies?.active?{returnDate:automaticReplies.returnDate,note:automaticReplies.note}:null;
    const oofFromFallback=sageAbsence?null:c.oofAbsence?{...c.oofAbsence,note:automaticReplies?.note||null}:p.outOfOffice?{returnDate:null,note:automaticReplies?.note||null}:null;
    const oofAbsence=sageAbsence?null:automaticReplies?.configured?automaticRepliesAbsence:oofFromFallback;
    const absent=Boolean(sageAbsence||oofAbsence);
    return {id:u.id,name:u.displayName||"",initials:initials(u),email:u.mail||u.userPrincipalName||"",department:u.department||MISSING_DEPARTMENT,company:companyKey,companyName:u.companyName||"",role:u.jobTitle||"",officeLocation:u.officeLocation||"",resourceCapacity:companyKey===MEETING_ROOM_COMPANY?place?.capacity??null:undefined,resourceLocationDescription:companyKey===MEETING_ROOM_COMPANY?resourceLocationDescription||null:undefined,photoUrl:`/api/users/${u.id}/photo`,location:p.location,presence:p.presence,availability:p.availability,activity:p.activity,outOfOffice:p.outOfOffice,outOfOfficeNote:automaticReplies?.note||null,automaticRepliesStatus:automaticReplies?.status||"unavailable",automaticRepliesActive:Boolean(automaticReplies?.active),phone:absent?"free":p.phone,currentMeetingEnd:absent?null:c.currentMeetingEnd,currentMeetingShowAs:absent?null:c.currentMeetingShowAs,meetingReachable:absent?false:c.meetingReachable,nextMeeting:absent||c.currentMeetingEnd?null:c.nextMeeting,resourceSchedule:companyKey===MEETING_ROOM_COMPANY?c.schedule:undefined,resourceCalendarAvailable:companyKey===MEETING_ROOM_COMPANY?calendars.has(u.id):undefined,sageAbsence,oofAbsence};
  });
  return {success:true,generatedAt:new Date().toISOString(),group,count:people.length,people};
}
const teamsNotifications=createTeamsNotificationService({getMembers,presences,mapPresence,graph});
app.get("/api/health",(req,res)=>res.json({success:true,status:"ok",generatedAt:new Date().toISOString(),teamsNotifications:teamsNotifications.health()}));
app.get("/api/presence-board",async(req,res)=>{try{res.set("Cache-Control","no-store");res.json(await build())}catch(e){console.error(e);res.status(500).json({success:false,error:"Presence-Daten konnten nicht geladen werden.",details:e.message})}});
app.get("/api/teams/notifications/subscriptions",async(req,res)=>{try{res.set("Cache-Control","no-store");res.json({personIds:await teamsNotifications.getSubscriptions(req.get("authorization"))})}catch(error){res.status(error.status||500).json({error:error.status?error.message:"Teams subscriptions could not be loaded."})}});
app.put("/api/teams/notifications/subscriptions",async(req,res)=>{try{const activeSubscriptions=await teamsNotifications.replaceSubscriptions(req.get("authorization"),req.body?.personIds);res.json({success:true,activeSubscriptions})}catch(error){res.status(error.status||500).json({error:error.status?error.message:"Teams subscriptions could not be saved."})}});
app.get("/api/users/:userId/photo",async(req,res)=>{try{const response=await fetch(`${GRAPH}/users/${encodeURIComponent(req.params.userId)}/photo/$value`,{headers:{Authorization:`Bearer ${await token()}`}});if(!response.ok)return res.sendStatus(404);res.set("Content-Type",response.headers.get("content-type")||"image/jpeg");res.set("Cache-Control","private, max-age=86400");res.send(Buffer.from(await response.arrayBuffer()))}catch{return res.sendStatus(404)}});
app.use((req,res)=>res.status(404).json({success:false,error:"API-Endpunkt nicht gefunden.",path:req.originalUrl}));
app.listen(PORT,()=>console.log(`Crew Presence Board API läuft auf Port ${PORT}`));
teamsNotifications.start();
