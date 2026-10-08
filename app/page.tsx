"use client";
import { useEffect, useState } from "react";
import { read, utils, writeFile } from "xlsx";
import { CalendarDays, Download, Upload, WandSparkles, Printer } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";

type Course = { code:string; title:string; groups:string[]; students:number; hours:number };
type Room = { name:string; capacity:number };
type Exam = Course & { date:string; start:string; end:string; room:string };
const examples:Course[] = [
  {code:"MTH 101",title:"General Mathematics I",groups:["Mathematics 100","Physics 100"],students:110,hours:2},
  {code:"PHY 101",title:"General Physics I",groups:["Physics 100","Chemistry 100"],students:90,hours:2},
  {code:"CHM 101",title:"General Chemistry I",groups:["Chemistry 100","Biochemistry 100"],students:78,hours:2},
  {code:"MTH 201",title:"Linear Algebra",groups:["Mathematics 200"],students:64,hours:2},
  {code:"STA 201",title:"Probability and Statistics",groups:["Mathematics 200","Statistics 200"],students:95,hours:2},
  {code:"CSC 201",title:"Programming Fundamentals",groups:["Statistics 200","Computer Science 200"],students:100,hours:2},
];
const exampleRooms:Room[]=[{name:"Faculty Hall",capacity:150},{name:"Lecture Theatre A",capacity:100},{name:"Lecture Theatre B",capacity:80}];
const min=(t:string)=>{const [h,m]=t.split(":").map(Number);return h*60+m};
const overlap=(a:{start:string;end:string},b:{start:string;end:string})=>min(a.start)<min(b.end)&&min(b.start)<min(a.end);
const cell=(v:unknown)=>String(v??"").trim();
function create(courses:Course[],rooms:Room[],dates:string[],sessions:{start:string;end:string}[],assignVenues:boolean){
  const exams:Exam[]=[];const unscheduled:string[]=[];
  for(const c of [...courses].sort((a,b)=>b.groups.length-a.groups.length||b.students-a.students||a.code.localeCompare(b.code))){
    let chosen:Exam|undefined;
    for(const date of dates){for(const s of sessions){if(min(s.end)-min(s.start)<c.hours*60)continue;
      const concurrent=exams.filter(e=>e.date===date&&overlap(e,s));
      if(concurrent.some(e=>e.groups.some(g=>c.groups.includes(g))))continue;
      if(!assignVenues){chosen={...c,date,...s,room:""};break}
      const free=[...rooms].filter(r=>!concurrent.some(e=>e.room.split("; ").includes(r.name))).sort((a,b)=>b.capacity-a.capacity);
      const selected:Room[]=[];let seats=0;
      for(const r of free){selected.push(r);seats+=r.capacity;if(seats>=c.students)break}
      if(seats>=c.students){chosen={...c,date,...s,room:selected.map(r=>r.name).join("; ")};break}
    }if(chosen)break;}
    if(chosen)exams.push(chosen);else unscheduled.push(c.code);
  }
  exams.sort((a,b)=>a.date.localeCompare(b.date)||a.start.localeCompare(b.start)||a.room.localeCompare(b.room));
  return {exams,unscheduled};
}
export default function Home(){
  const [courses,setCourses]=useState(examples),[rooms,setRooms]=useState(exampleRooms);
  const [date,setDate]=useState("2026-11-02"),[days,setDays]=useState(6),[sessions,setSessions]=useState("08:00-10:30, 11:00-13:30, 14:00-17:00");
  const [includeSaturday,setIncludeSaturday]=useState(true),[assignVenues,setAssignVenues]=useState(false);
  const [file,setFile]=useState(""),[message,setMessage]=useState(""),[result,setResult]=useState<ReturnType<typeof create>|null>(null);
  useEffect(()=>{
    const context=(document as Document & {modelContext?:{registerTool:(tool:unknown,options:{signal:AbortSignal})=>Promise<void>|void}}).modelContext;
    if(!context?.registerTool)return;
    const controller=new AbortController();
    Promise.resolve(context.registerTool({name:"generate_exam_timetable",title:"Generate exam timetable",description:"Generate and display a timetable from the currently imported courses, rooms, dates and sessions.",inputSchema:{type:"object",properties:{},additionalProperties:false},annotations:{readOnlyHint:false,untrustedContentHint:false},execute:()=>{generate();return {status:"generation_requested",courseCount:courses.length,roomCount:rooms.length}}},{signal:controller.signal})).catch(()=>{});
    return ()=>controller.abort();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[courses,rooms,date,days,sessions,includeSaturday,assignVenues]);
  function template(){const b=utils.book_new();utils.book_append_sheet(b,utils.json_to_sheet(examples.map(c=>({Code:c.code,Title:c.title,Groups:c.groups.join("; "),Students:c.students,"Duration (hours)":c.hours}))),"Courses");utils.book_append_sheet(b,utils.json_to_sheet(exampleRooms.map(r=>({Room:r.name,Capacity:r.capacity}))),"Rooms");writeFile(b,"faculty-exams-template.xlsx")}
  async function importWorkbook(f?:File){if(!f)return;try{
    const b=read(await f.arrayBuffer(),{type:"array"});if(!b.Sheets.Courses||!b.Sheets.Rooms)throw Error("Workbook needs Courses and Rooms sheets. Download the template for the format.");
    const cs=utils.sheet_to_json<Record<string,unknown>>(b.Sheets.Courses).map(r=>({code:cell(r.Code),title:cell(r.Title),groups:cell(r.Groups).split(/[;,]/).map(x=>x.trim()).filter(Boolean),students:Number(r.Students),hours:Number(r["Duration (hours)"])}));
    const rs=utils.sheet_to_json<Record<string,unknown>>(b.Sheets.Rooms).map(r=>({name:cell(r.Room),capacity:Number(r.Capacity)}));
    if(!cs.length||!rs.length||cs.some(c=>!c.code||!c.title||!c.groups.length||!Number.isFinite(c.students)||c.students<=0||!Number.isFinite(c.hours)||c.hours<=0)||rs.some(r=>!r.name||!Number.isFinite(r.capacity)||r.capacity<=0))throw Error("Fill every field with valid course, group, room and capacity data.");
    if(new Set(cs.map(c=>c.code)).size!==cs.length||new Set(rs.map(r=>r.name)).size!==rs.length)throw Error("Course codes and room names must be unique.");
    setCourses(cs);setRooms(rs);setFile(f.name);setResult(null);setMessage(`${cs.length} courses and ${rs.length} rooms imported.`);
  }catch(e){setMessage(e instanceof Error?e.message:"Could not read workbook.")}}
  function generate(){
    const ss=sessions.split(",").map(s=>s.trim()).filter(Boolean).map(s=>s.match(/^(\d{2}:\d{2})-(\d{2}:\d{2})$/));
    if(!date||days<1||days>60||!ss.length||ss.some(m=>!m||min(m[1])>=min(m[2])||min(m[2])>1440)){setMessage("Enter a date, 1–60 days, and sessions such as 09:00-12:00, 13:00-16:00.");return}
    const dates=Array.from({length:days},(_,i)=>{const d=new Date(date+"T12:00:00Z");d.setUTCDate(d.getUTCDate()+i);return d}).filter(d=>d.getUTCDay()!==0&&(includeSaturday||d.getUTCDay()!==6)).map(d=>d.toISOString().slice(0,10));
    if(!dates.length){setMessage("The selected period has no exam days.");return}
    setResult(create(courses,rooms,dates,ss.map(m=>({start:m![1],end:m![2]})),assignVenues));setMessage("");
  }
  const sessionList=sessions.split(",").map(s=>s.trim()).filter(Boolean).map(s=>s.match(/^(\d{2}:\d{2})-(\d{2}:\d{2})$/)).filter((m):m is RegExpMatchArray=>!!m).map(m=>({start:m[1],end:m[2]}));
  const examDates=result?[...new Set(result.exams.map(e=>e.date))]:[];
  function exportWorkbook(){if(!result)return;const b=utils.book_new();
    const layout:(string|number)[][]=[];const merges:{s:{r:number;c:number};e:{r:number;c:number}}[]=[];
    for(const date of examDates){const headerRow=layout.length;layout.push([new Date(date+"T12:00:00Z").toLocaleDateString("en-GB",{weekday:"long",day:"numeric",month:"long",year:"numeric",timeZone:"UTC"}).toUpperCase()]);merges.push({s:{r:headerRow,c:0},e:{r:headerRow,c:Math.max(1,sessionList.length*2-1)}});
      layout.push(sessionList.flatMap(s=>[s.start+" – "+s.end,""]));layout.push(sessionList.flatMap(()=>["COURSE(S)","VENUE(S)"]));
      const columns=sessionList.map(s=>result.exams.filter(e=>e.date===date&&e.start===s.start&&e.end===s.end));
      for(let i=0;i<Math.max(...columns.map(x=>x.length),0);i++)layout.push(columns.flatMap(col=>col[i]?[col[i].code+" ("+col[i].students+")",col[i].room]:["",""]));
      layout.push([]);
    }
    const sheet=utils.aoa_to_sheet(layout);sheet["!merges"]=merges;sheet["!cols"]=sessionList.flatMap(()=>[{wch:24},{wch:22}]);utils.book_append_sheet(b,sheet,"Faculty layout");
    utils.book_append_sheet(b,utils.json_to_sheet(result.exams.map(e=>({Date:e.date,Start:e.start,End:e.end,Code:e.code,Course:e.title,Groups:e.groups.join("; "),Students:e.students,Venue:e.room}))),"Exam details");
    utils.book_append_sheet(b,utils.json_to_sheet(result.unscheduled.map(code=>({Code:code})),{header:["Code"]}),"Unscheduled");writeFile(b,"faculty-exam-timetable.xlsx")}
  return <main className="shell"><header><div className="brand"><span><CalendarDays size={23}/></span> ExamGrid <small>FACULTY PLANNER</small></div><div className="header-label">Examination timetable workspace</div></header>
  <div className="content"><div className="intro"><div><p className="eyebrow">PLAN · CHECK · EXPORT</p><h1>Examination timetable</h1><p>Import courses and rooms, choose the exam period, and produce a timetable without student group or room clashes.</p></div><span className="badge">Working draft</span></div>
  <div className="grid"><section className="panel"><div className="heading"><b>01</b><div><h2>Course & room data</h2><p>Use the sample or import an Excel workbook.</p></div></div><label className="upload"><Upload size={24}/><strong>{file||"Choose .xlsx workbook"}</strong><span>Courses and Rooms sheets</span><input type="file" accept=".xlsx,.xls" onChange={e=>importWorkbook(e.target.files?.[0])}/></label><button className="link" onClick={template}><Download size={17}/> Download Excel template</button><div className="stats"><div><strong>{courses.length}</strong><span>Courses</span></div><div><strong>{rooms.length}</strong><span>Rooms</span></div><div><strong>{new Set(courses.flatMap(c=>c.groups)).size}</strong><span>Groups</span></div></div><p className="help">Use identical group names for courses taken by the same students. Separate multiple groups with semicolons.</p></section>
  <section className="panel"><div className="heading"><b>02</b><div><h2>Exam period</h2><p>Set exam days and the three faculty sessions.</p></div></div><div className="fields"><label>First exam date<Input type="date" value={date} onChange={e=>setDate(e.target.value)}/></label><label>Calendar days<Input type="number" min="1" max="60" value={days} onChange={e=>setDays(Number(e.target.value))}/></label></div><label className="field">Daily sessions <span>24-hour format, comma separated</span><Input value={sessions} onChange={e=>setSessions(e.target.value)}/></label><div className="options"><label><input type="checkbox" checked={includeSaturday} onChange={e=>setIncludeSaturday(e.target.checked)}/> Include Saturdays</label><label><input type="checkbox" checked={assignVenues} onChange={e=>setAssignVenues(e.target.checked)}/> Assign venues from Rooms sheet</label></div><p className="help">{assignVenues?"Venues are combined when one room cannot hold all students.":"Venue columns stay blank, matching the first draft you shared."} Sundays are skipped.</p><Button className="generate" onClick={generate}><WandSparkles size={18}/> Generate timetable</Button>{message&&<p role="status" className="notice">{message}</p>}</section></div>
  <section className="panel results"><div className="result-head"><div className="heading"><b>03</b><div><h2>Faculty timetable</h2><p>{result?`${result.exams.length} of ${courses.length} exams scheduled`:"Generate a timetable to review assignments."}</p></div></div>{result&&<div className="result-actions"><Button variant="outline" onClick={()=>window.print()}><Printer size={17}/> Print</Button><Button variant="outline" onClick={exportWorkbook}><Download size={17}/> Export Excel</Button></div>}</div>{result?<><div className={result.unscheduled.length?"status warning":"status good"}>{result.unscheduled.length?`${result.unscheduled.length} unscheduled: ${result.unscheduled.join(", ")}. Add days, sessions or venues.`:"All courses placed. No student group has overlapping exams."}</div>
  <div className="faculty-layout">{examDates.map(d=><div className="day-block" key={d}><h3>{new Date(d+"T12:00:00Z").toLocaleDateString("en-GB",{weekday:"long",day:"numeric",month:"long",year:"numeric",timeZone:"UTC"})}</h3><div className="session-grid">{sessionList.map(s=><div className="session" key={s.start+"-"+s.end}><h4>{s.start} – {s.end}</h4><div className="session-columns"><span>COURSE(S)</span><span>VENUE(S)</span></div>{result.exams.filter(e=>e.date===d&&e.start===s.start&&e.end===s.end).map(e=><div className="exam-line" key={e.code}><span>{e.code} ({e.students})</span><span>{e.room||"—"}</span></div>)}</div>)}</div></div>)}</div>
  <details className="details"><summary>View full course details</summary><div className="table-wrap"><table><thead><tr><th>Date</th><th>Time</th><th>Course</th><th>Student groups</th><th>Students</th><th>Venue</th></tr></thead><tbody>{result.exams.map(e=><tr key={e.code}><td>{new Date(e.date+"T12:00:00Z").toLocaleDateString("en-GB",{weekday:"short",day:"numeric",month:"short",timeZone:"UTC"})}</td><td>{e.start}–{e.end}</td><td><strong>{e.code}</strong><small>{e.title}</small></td><td>{e.groups.join(", ")}</td><td>{e.students}</td><td>{e.room||"Pending"}</td></tr>)}</tbody></table></div></details></>:<div className="empty"><CalendarDays size={34}/><strong>Your timetable will appear here</strong><span>Sample data is ready. Choose Generate timetable.</span></div>}</section>
  <footer>Review this draft with the faculty exam office before publication. Confirm registration and room availability.</footer></div></main>
}
