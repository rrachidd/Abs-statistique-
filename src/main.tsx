import './index.css';
import { auth, db, googleProvider, signInWithPopup, signOut, onAuthStateChanged, collection, doc, setDoc, getDocs, deleteDoc } from './firebase';
import { GoogleGenAI } from '@google/genai';
import { marked } from 'marked';

// Expose global variables and functions for inline HTML event handlers
const MAX_FILE_SIZE=5*1024*1024;
const MONTH_NAMES={'janvier':'يناير','février':'فبراير','mars':'مارس','avril':'أبريل','mai':'ماي','juin':'يونيو','juillet':'يوليوز','août':'غشت','septembre':'شتنبر','octobre':'أكتوبر','novembre':'نونبر','décembre':'دجنبر','يناير':'يناير','فبراير':'فبراير','مارس':'مارس','أبريل':'أبريل','ماي':'ماي','يونيو':'يونيو','يوليوز':'يوليوز','غشت':'غشت','شتنبر':'شتنبر','أكتوبر':'أكتوبر','نونبر':'نونبر','دجنبر':'دجنبر','january':'يناير','february':'فبراير','march':'مارس','april':'أبريل','may':'ماي','june':'يونيو','july':'يوليوز','august':'غشت','september':'شتنبر','october':'أكتوبر','november':'نونبر','december':'دجنبر'};
const MONTH_ORDER=['يناير','فبراير','مارس','أبريل','ماي','يونيو','يوليوز','غشت','شتنبر','أكتوبر','نونبر','دجنبر'];
const CHART_COLORS=['#059669','#D97706','#DC2626','#7C3AED','#2563EB','#0891B2','#C026D3','#EA580C','#65A30D','#0D9488','#9333EA','#E11D48','#CA8A04','#0369A1','#BE185D','#4F46E5','#15803D','#B45309'];

const state: any = {datasets:[],guardians:{},guardianDetails:{},activeClass:null,activeMonth:null,searchQuery:'',viewMode:'compact',sortMode:'rank',barChartMode:'total',latenessRecords:[]};
let barChart: any = null, lineChart: any = null, pieChart: any = null, hBarChart: any = null, currentSheetStudentId: any = null;
let currentUser: any = null;

function normCell(v: any){if(v===null||v===undefined||v==='')return'*';const s=String(v).trim();return(s==='*'||s===''||s.toLowerCase()==='null')?'*':s}
function parseDayValue(v: any){const s=normCell(v);if(s==='*')return{type:'none',val:0};if(s==='0')return{type:'present',val:0};if(s.toUpperCase()==='X')return{type:'special',val:0};const n=parseFloat(s);return(!isNaN(n)&&n>0)?{type:'absent',val:n}:{type:'present',val:0}}
function translateMonth(m: any){return(!m)?'غير محدد':(MONTH_NAMES[m.toLowerCase().trim() as keyof typeof MONTH_NAMES]||m)}
function getSchoolDays(students: any[]){const d=new Set();students.forEach(st=>{for(let i=1;i<=31;i++){if(st.days[i]!==undefined&&parseDayValue(st.days[i]).type!=='none')d.add(i)}});return d}
function calcTotalAbsences(st: any){let t=0;for(let d=1;d<=31;d++){const i=parseDayValue(st.days[d]);if(i.type==='absent')t+=i.val}return t}
function calcAbsentDays(st: any){let c=0;for(let d=1;d<=31;d++){if(parseDayValue(st.days[d]).type==='absent')c++}return c}
function getJustified(st: any){if(!st.summaries||st.summaries.length===0)return 0;const d=parseInt(st.summaries[1])||0;const h=parseInt(st.summaries[3])||0;const total=calcTotalAbsences(st);return Math.min((d*4)+h, total)}
function getUnjustified(st: any){const total=calcTotalAbsences(st);const just=getJustified(st);return Math.max(0,total-just)}
function studentFullName(st: any){return`${st.family} ${st.name}`}
function getClassKey(ds: any){return ds.metadata.class||'غير محدد'}
function extractAfter(text: string,keyword: string){const idx=text.indexOf(keyword);if(idx===-1)return'';let rest=text.substring(idx+keyword.length).replace(/[:\s]+/,' ').trim();['المديرية','نيابة','عمالة','السنة الدراسية','المستوى','القسم','الشهر','غير مبرر','مبرر'].forEach(nk=>{const ni=rest.indexOf(nk);if(ni>0)rest=rest.substring(0,ni).trim()});return rest.replace(/[*\|]/g,'').trim()}

function getUniqueClasses(){const map=new Map();state.datasets.forEach((ds: any)=>{const k=getClassKey(ds);if(!map.has(k))map.set(k,{institution:ds.metadata.institution,level:ds.metadata.level,year:ds.metadata.year,academy:ds.metadata.academy})});return map}
function getClassMonths(cn: string){return state.datasets.filter((ds: any)=>getClassKey(ds)===cn).sort((a: any,b: any)=>MONTH_ORDER.indexOf(a.metadata.monthAr)-MONTH_ORDER.indexOf(b.metadata.monthAr))}
function getActiveDataset(){return state.datasets.find((ds: any)=>getClassKey(ds)===state.activeClass&&ds.metadata.month===state.activeMonth)}
function getActiveClassDatasets(){return state.datasets.filter((ds: any)=>getClassKey(ds)===state.activeClass).sort((a: any,b: any)=>MONTH_ORDER.indexOf(a.metadata.monthAr)-MONTH_ORDER.indexOf(b.metadata.monthAr))}

function getDatasetId(ds: any) {
    let id = `${ds.metadata.class}_${ds.metadata.month}`.replace(/\//g, '-');
    if (id === '.' || id === '..') id = `ds_${id}`;
    if (id.startsWith('__') && id.endsWith('__')) id = `ds_${id}`;
    return id;
}

function showToast(msg: string,type='info',dur=3500){const c=document.getElementById('toast-container');if(!c)return;const icons: any={success:'fa-circle-check',error:'fa-circle-xmark',info:'fa-circle-info'};const t=document.createElement('div');t.className=`toast toast-${type}`;t.innerHTML=`<i class="fa-solid ${icons[type]||icons.info}"></i><span>${msg}</span>`;c.appendChild(t);setTimeout(()=>{t.classList.add('removing');setTimeout(()=>t.remove(),400)},dur)}

function saveDataLocally() {
    try {
        localStorage.setItem('absence_datasets', JSON.stringify(state.datasets));
        localStorage.setItem('absence_guardians', JSON.stringify(state.guardians));
        localStorage.setItem('absence_guardianDetails', JSON.stringify(state.guardianDetails));
        localStorage.setItem('absence_lateness', JSON.stringify(state.latenessRecords));
    } catch (e) {
        console.error("Local storage error", e);
    }
}

function loadDataLocally() {
    try {
        const ds = localStorage.getItem('absence_datasets');
        const gd = localStorage.getItem('absence_guardians');
        const gdd = localStorage.getItem('absence_guardianDetails');
        const lat = localStorage.getItem('absence_lateness');
        if (ds) state.datasets = JSON.parse(ds);
        if (gd) state.guardians = JSON.parse(gd);
        if (gdd) state.guardianDetails = JSON.parse(gdd);
        if (lat) state.latenessRecords = JSON.parse(lat);
        
        if (state.datasets.length > 0) {
            if (!state.activeClass || !state.datasets.find((d: any)=>getClassKey(d)===state.activeClass)) {
                state.activeClass = getClassKey(state.datasets[0]);
            }
            const cm = getClassMonths(state.activeClass);
            if (!cm.find((d: any)=>d.metadata.month===state.activeMonth)) {
                state.activeMonth = cm.length ? cm[0].metadata.month : null;
            }
        }
    } catch (e) {
        console.error("Local storage load error", e);
    }
}

async function saveDatasetToFirebase(dataset: any) {
    saveDataLocally();
    if (!currentUser) return;
    try {
        const datasetId = getDatasetId(dataset);
        const docRef = doc(db, `users/${currentUser.uid}/datasets`, datasetId);
        await setDoc(docRef, {
            userId: currentUser.uid,
            fileName: dataset.fileName || '',
            metadata: dataset.metadata,
            students: dataset.students,
            summaryCols: dataset.summaryCols || [],
            createdAt: Date.now()
        });
        console.log("Dataset saved to Firebase:", datasetId);
    } catch (error) {
        console.error("Error saving dataset:", error);
        showToast('حدث خطأ أثناء الحفظ في قاعدة البيانات', 'error');
    }
}

async function saveGuardiansToFirebase() {
    saveDataLocally();
    if (!currentUser) return;
    try {
        const docRef = doc(db, `users/${currentUser.uid}/settings`, 'guardians');
        await setDoc(docRef, { data: state.guardians, details: state.guardianDetails });
        console.log("Guardians saved to Firebase");
    } catch (error) {
        console.error("Error saving guardians:", error);
    }
}

async function loadGuardiansFromFirebase() {
    if (!currentUser) return;
    try {
        const querySnapshot = await getDocs(collection(db, `users/${currentUser.uid}/settings`));
        querySnapshot.forEach((doc) => {
            if (doc.id === 'guardians') {
                state.guardians = doc.data().data || {};
                state.guardianDetails = doc.data().details || {};
            }
        });
    } catch (error) {
        console.error("Error loading guardians:", error);
    }
}

async function loadDatasetsFromFirebase() {
    if (!currentUser) return;
    try {
        const querySnapshot = await getDocs(collection(db, `users/${currentUser.uid}/datasets`));
        const datasets: any[] = [];
        querySnapshot.forEach((doc) => {
            datasets.push(doc.data());
        });
        state.datasets = datasets;
        
        if (state.datasets.length > 0) {
            if (!state.activeClass || !state.datasets.find((d: any)=>getClassKey(d)===state.activeClass)) {
                state.activeClass = getClassKey(state.datasets[0]);
            }
            const cm = getClassMonths(state.activeClass);
            if (!cm.find((d: any)=>d.metadata.month===state.activeMonth)) {
                state.activeMonth = cm.length ? cm[0].metadata.month : null;
            }
        }
        renderAll();
        showToast('تم استرجاع البيانات بنجاح', 'success');
    } catch (error) {
        console.error("Error loading datasets:", error);
        showToast('حدث خطأ أثناء استرجاع البيانات', 'error');
    }
}

async function deleteDatasetFromFirebase(datasetId: string) {
    if (!currentUser) return;
    try {
        await deleteDoc(doc(db, `users/${currentUser.uid}/datasets`, datasetId));
    } catch (error) {
        console.error("Error deleting dataset:", error);
    }
}

function processExcel(buffer: any,fileName: string){
    try{
        const wb=(window as any).XLSX.read(buffer,{type:'array'});const sn=wb.SheetNames.includes('Data')?'Data':wb.SheetNames[0];
        const raw=(window as any).XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:'',blankrows:false});
        if(raw.length<6){showToast(`"${fileName}" فارغ`,'error');return}
        const parsed=parseData(raw);if(!parsed||parsed.students.length===0){showToast(`لا بيانات في "${fileName}"`,'error');return}
        const ck=getClassKey(parsed);const ex=state.datasets.findIndex((d: any)=>getClassKey(d)===ck&&d.metadata.month===parsed.metadata.month);
        
        const newDataset = {fileName,...parsed};
        if(ex!==-1){
            state.datasets[ex]={...state.datasets[ex], ...newDataset};
            showToast(`تم تحديث ${parsed.metadata.monthAr} — ${parsed.metadata.class}`,'info')
        } else {
            state.datasets.push(newDataset);
            showToast(`تم استيراد ${parsed.metadata.monthAr} — ${parsed.metadata.class}`,'success')
        }
        
        if(!state.activeClass||!state.datasets.find((d: any)=>getClassKey(d)===state.activeClass))state.activeClass=ck;
        const cm=getClassMonths(state.activeClass);if(!cm.find((d: any)=>d.metadata.month===state.activeMonth))state.activeMonth=cm.length?cm[0].metadata.month:null;
        
        renderAll();
        
        if (currentUser) {
            saveDatasetToFirebase(newDataset);
        } else {
            showToast('البيانات محفوظة محلياً فقط. يرجى تسجيل الدخول لحفظها في السحابة.', 'info', 5000);
        }
    }catch(err: any){console.error(err);showToast(`خطأ: ${err.message}`,'error')}
}

function processGuardianExcel(buffer: any) {
    try {
        const wb = (window as any).XLSX.read(buffer, { type: 'array' });
        const sn = wb.SheetNames[0];
        const raw = (window as any).XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: '', blankrows: false });
        
        if (raw.length < 2) {
            showToast(`ملف أولياء الأمور فارغ`, 'error');
            return;
        }

        // Find header row
        let headerIdx = -1;
        let massarCol = -1;
        let nameCol = -1;
        let familyCol = -1;
        let fullNameCol = -1;
        let phoneCols: number[] = [];
        let fatherCol = -1;
        let motherCol = -1;
        let guardianCol = -1;

        for (let i = 0; i < Math.min(10, raw.length); i++) {
            const row = raw[i] as string[];
            for (let j = 0; j < row.length; j++) {
                const cell = String(row[j]).trim().toLowerCase();
                if (cell.includes('مسار') || cell.includes('رقم التلميذ') || cell.includes('الرقم الوطني') || cell.includes('massar')) massarCol = j;
                if (cell.includes('هاتف') || cell.includes('رقم الهاتف') || cell.includes('الهاتف') || cell.includes('phone') || cell.includes('téléphone') || cell.includes('tel') || cell.includes('أب') || cell.includes('أم') || cell.includes('ولي')) {
                    if (cell.includes('هاتف') || cell.includes('tel') || cell.includes('phone') || cell.includes('gsm') || cell.includes('محمول')) {
                        if (!phoneCols.includes(j)) phoneCols.push(j);
                    }
                }
                if (cell === 'الاسم' || cell === 'الإسم' || cell.includes('الإسم بالعربية') || cell.includes('الاسم بالعربية') || cell === 'prenom' || cell === 'prénom') nameCol = j;
                if (cell === 'النسب' || cell.includes('النسب بالعربية') || cell === 'nom') familyCol = j;
                if (cell.includes('الاسم الكامل') || cell.includes('الإسم الكامل') || cell === 'التلميذ' || cell === 'اسم التلميذ' || cell === 'nom et prenom' || cell === 'nom complet') fullNameCol = j;
                
                if (cell.includes('اسم الأب') || cell.includes('إسم الأب') || cell.includes('الأب') || cell.includes('pere') || cell.includes('père')) fatherCol = j;
                if (cell.includes('اسم الأم') || cell.includes('إسم الأم') || cell.includes('الأم') || cell.includes('mere') || cell.includes('mère')) motherCol = j;
                if (cell.includes('اسم الولي') || cell.includes('إسم الولي') || cell.includes('الولي') || cell.includes('tuteur')) guardianCol = j;
            }
            if (massarCol !== -1 || fullNameCol !== -1 || (nameCol !== -1 && familyCol !== -1)) {
                headerIdx = i;
                break;
            }
        }

        if (headerIdx === -1) {
            showToast(`لم يتم العثور على أعمدة تعريف التلميذ في الملف`, 'error');
            return;
        }

        let addedCount = 0;
        for (let i = headerIdx + 1; i < raw.length; i++) {
            const row = raw[i] as string[];
            
            // Extract all phone numbers from the row
            let phones: string[] = [];
            
            // First check phone columns
            phoneCols.forEach(col => {
                const val = String(row[col] || '').trim();
                if (val) {
                    const normalizedVal = val.replace(/[\s\.\-]/g, '');
                    const matches = normalizedVal.match(/(?:\+|00)?(?:212|0)[567]\d{8}/g) || normalizedVal.match(/\+?\d{9,15}/g);
                    if (matches) phones.push(...matches);
                }
            });

            // If no phones found in phone columns, scan the whole row
            if (phones.length === 0) {
                for (let j = 0; j < row.length; j++) {
                    if (j === massarCol || j === nameCol || j === familyCol || j === fullNameCol) continue;
                    const val = String(row[j] || '').trim();
                    if (val) {
                        const normalizedVal = val.replace(/[\s\.\-]/g, '');
                        const matches = normalizedVal.match(/(?:\+|00)?(?:212|0)[567]\d{8}/g) || normalizedVal.match(/\+?\d{9,15}/g);
                        if (matches) phones.push(...matches);
                    }
                }
            }

            // Deduplicate phones
            phones = [...new Set(phones)];
            const phoneStr = phones.join(' - ');
            
            const fatherName = fatherCol !== -1 ? String(row[fatherCol] || '').trim() : '';
            const motherName = motherCol !== -1 ? String(row[motherCol] || '').trim() : '';
            const guardianName = guardianCol !== -1 ? String(row[guardianCol] || '').trim() : '';

            if (!phoneStr && !fatherName && !motherName && !guardianName) continue;

            let keys = [];
            if (massarCol !== -1) {
                const massar = String(row[massarCol] || '').trim().toLowerCase();
                if (massar) keys.push(massar);
            }
            
            let family = '';
            let name = '';
            if (familyCol !== -1) family = String(row[familyCol] || '').trim().toLowerCase();
            if (nameCol !== -1) name = String(row[nameCol] || '').trim().toLowerCase();
            
            if (family && name) {
                keys.push(`${family} ${name}`.replace(/\s+/g, ' '));
                keys.push(`${name} ${family}`.replace(/\s+/g, ' '));
            }
            
            if (fullNameCol !== -1) {
                const fullName = String(row[fullNameCol] || '').trim().toLowerCase().replace(/\s+/g, ' ');
                if (fullName) keys.push(fullName);
            }

            if (keys.length > 0) {
                keys.forEach(k => {
                    if (phoneStr) {
                        if (state.guardians[k] && state.guardians[k] !== phoneStr) {
                            const existing = state.guardians[k].split(' - ');
                            const combined = [...new Set([...existing, ...phones])].join(' - ');
                            state.guardians[k] = combined;
                        } else {
                            state.guardians[k] = phoneStr;
                        }
                    }
                    
                    if (!state.guardianDetails[k]) state.guardianDetails[k] = {};
                    if (fatherName) state.guardianDetails[k].father = fatherName;
                    if (motherName) state.guardianDetails[k].mother = motherName;
                    if (guardianName) state.guardianDetails[k].guardian = guardianName;
                });
                addedCount++;
            }
        }

        if (addedCount > 0) {
            saveGuardiansToFirebase();
            showToast(`تم استيراد وتحيين أرقام هواتف ${addedCount} ولي أمر بنجاح`, 'success');
            renderTable();
            // Re-render WhatsApp modal if it's open
            const m = document.getElementById('whatsapp-modal');
            if (m && m.classList.contains('open')) {
                (window as any).openWhatsAppModal();
            }
        } else {
            showToast(`لم يتم العثور على أرقام هواتف صالحة في الملف`, 'info');
        }

    } catch (err: any) {
        console.error(err);
        showToast(`خطأ في استيراد ملف أولياء الأمور: ${err.message}`, 'error');
    }
}

function parseData(raw: any[]){
    let headerIdx=-1,subHeaderIdx=-1;
    for(let i=0;i<Math.min(12,raw.length);i++){const row=raw[i]||[];if(row.some((c: any)=>String(c).includes('الترتيب')))headerIdx=i;if(row.some((c: any)=>String(c).includes('يوم'))&&row.some((c: any)=>String(c).includes('ساعة')))subHeaderIdx=i}
    if(headerIdx===-1)return null;
    const headers=raw[headerIdx]||[];const colMap: any={dayCols:{}};
    headers.forEach((h: any,i: number)=>{const s=String(h).trim();if(s==='الترتيب')colMap.rank=i;else if(s==='رقم التلميذ')colMap.id=i;else if(s==='النسب بالعربية')colMap.family=i;else if(s==='الإسم بالعربية')colMap.name=i;else if(s==='المجموع')colMap.total=i;else{const n=parseInt(s);if(!isNaN(n)&&n>=1&&n<=31)colMap.dayCols[n]=i}});
    const meta={month:'',monthAr:'',academy:'',institution:'',level:'',class:'',year:''};
    for(let i=0;i<headerIdx;i++){
        const row=raw[i]||[];
        const text=row.join(' ');
        if(text.includes('أكاديمية'))meta.academy=extractAfter(text,'أكاديمية');
        if(text.includes('مؤسسة'))meta.institution=extractAfter(text,'مؤسسة');
        if(text.includes('المستوى'))meta.level=extractAfter(text,'المستوى');
        if(text.includes('القسم'))meta.class=extractAfter(text,'القسم');
        if(text.includes('الشهر')){const m=extractAfter(text,'الشهر');meta.month=m.toLowerCase();meta.monthAr=translateMonth(m)}
        if(text.includes('السنة')){const y=text.match(/\d{4}\s*\/\s*\d{4}/);if(y)meta.year=y[0]}
    }
    const summaryCols: any[]=[];
    if(subHeaderIdx!==-1){const subRow=raw[subHeaderIdx]||[];for(let i=(colMap.total!==undefined?colMap.total:35);i<subRow.length;i++){const v=String(subRow[i]).trim();if(['يوم','ساعة','X','1/2'].includes(v))summaryCols.push({idx:i,label:v})}}
    const startRow=(subHeaderIdx!==-1?subHeaderIdx:headerIdx)+1;const students=[];
    for(let i=startRow;i<raw.length;i++){const row=raw[i]||[];const rank=row[colMap.rank],id=row[colMap.id];if(!rank&&!id)continue;const st: any={rank:String(rank).trim(),id:String(id).trim(),family:String(row[colMap.family]||'').trim(),name:String(row[colMap.name]||'').trim(),days:{},totalRaw:row[colMap.total]!==undefined?String(row[colMap.total]).trim():'',summaries:[]};for(let d=1;d<=31;d++){const ci=colMap.dayCols[d];if(ci!==undefined)st.days[d]=normCell(row[ci])}summaryCols.forEach(sc=>st.summaries.push(String(row[sc.idx]||'').trim()));students.push(st)}
    return{metadata:meta,students,summaryCols};
}

function renderAll(){
    const has=state.datasets.length>0;
    const uploadSection = document.getElementById('upload-section');
    const dataSection = document.getElementById('data-section');
    const btnPrint = document.getElementById('btn-print');
    const btnPdf = document.getElementById('btn-pdf');
    
    if(uploadSection) uploadSection.style.display=has?'none':'';
    if(dataSection) dataSection.style.display=has?'':'none';
    if(btnPrint) btnPrint.style.display=has?'':'none';
    if(btnPdf) btnPdf.style.display=has?'':'none';
    
    if(!has)return;
    renderClassTabs();renderMonthTabs();renderClassInfo();renderStats();renderCharts();renderTable();
}

function renderClassTabs(){const classes=getUniqueClasses();const c=document.getElementById('class-tabs');if(!c)return;c.innerHTML='';classes.forEach((info,cn)=>{const months=getClassMonths(cn);const btn=document.createElement('button');btn.className=`class-tab ${cn===state.activeClass?'active':''}`;btn.innerHTML=`<i class="fa-solid fa-users-rectangle text-sm ${cn===state.activeClass?'text-emerald-600':'text-gray-400'}"></i><span>${cn}</span><span class="month-count">${months.length} ${months.length===1?'شهر':'أشهر'}</span>`;btn.onclick=()=>{state.activeClass=cn;const cm=getClassMonths(cn);if(!cm.find((d: any)=>d.metadata.month===state.activeMonth))state.activeMonth=cm.length?cm[0].metadata.month:null;state.searchQuery='';(document.getElementById('search-input') as HTMLInputElement).value='';renderAll()};c.appendChild(btn)});const tc = document.getElementById('tabs-connector'); if(tc) tc.style.display=classes.size>0?'':'none'}
function renderMonthTabs(){const c=document.getElementById('month-tabs');const w=document.getElementById('month-tabs-container');if(!w || !c)return;if(!state.activeClass){w.style.display='none';return}const months=getClassMonths(state.activeClass);if(!months.length){w.style.display='none';return}w.style.display='';c.innerHTML='';months.forEach((ds: any)=>{const btn=document.createElement('button');btn.className=`month-tab ${ds.metadata.month===state.activeMonth?'active':''}`;btn.textContent=ds.metadata.monthAr||'—';btn.onclick=()=>{state.activeMonth=ds.metadata.month;state.searchQuery='';(document.getElementById('search-input') as HTMLInputElement).value='';renderAll()};c.appendChild(btn)})}
function renderClassInfo(){const ds=getActiveDataset();const ci=getUniqueClasses().get(state.activeClass);if(!ds&&!ci)return;const m=ds?ds.metadata:{};
    const cnd = document.getElementById('class-name-display'); if(cnd) cnd.textContent=state.activeClass;
    const bi = document.getElementById('bar-institution'); if(bi && bi.querySelector('span')) bi.querySelector('span')!.textContent=ci?.institution||m.institution||'—';
    const bl = document.getElementById('bar-level'); if(bl && bl.querySelector('span')) bl.querySelector('span')!.textContent=ci?.level||m.level||'—';
    const by = document.getElementById('bar-year'); if(by && by.querySelector('span')) by.querySelector('span')!.textContent=ci?.year||m.year||'—';
    const bs = document.getElementById('bar-students'); if(bs && bs.querySelector('span')) bs.querySelector('span')!.textContent=ds?`${ds.students.length} تلميذ`:'—';
}

function renderStats(){const ds=getActiveDataset();if(!ds)return;const students=ds.students;const sd=getSchoolDays(students);let ta=0;let ma={name:'—',val:0};students.forEach((st: any)=>{const a=calcTotalAbsences(st);ta+=a;if(a>ma.val)ma={name:studentFullName(st),val:a}});const stats=[{icon:'fa-user-group',bg:'bg-emerald-50',ic:'text-emerald-600',nc:'text-emerald-700',label:'عدد التلاميذ',value:students.length},{icon:'fa-clock',bg:'bg-red-50',ic:'text-red-500',nc:'text-red-600',label:'مجموع الغيابات',value:ta},{icon:'fa-triangle-exclamation',bg:'bg-amber-50',ic:'text-amber-600',nc:'text-amber-700',label:'أكثر غياباً',value:ma.name,sub:`${ma.val} غياب`},{icon:'fa-calendar-day',bg:'bg-teal-50',ic:'text-teal-600',nc:'text-teal-700',label:'أيام الدراسة',value:sd.size,sub:'من 31 يوم'}];
    const sc = document.getElementById('stats-cards');
    if(sc) sc.innerHTML=stats.map((s,i)=>`<div class="card card-hover p-5 fade-up" style="animation-delay:${i*80}ms"><div class="w-10 h-10 rounded-xl ${s.bg} flex items-center justify-center mb-3"><i class="fa-solid ${s.icon} ${s.ic}"></i></div><p class="text-2xl font-extrabold ${s.nc} mb-0.5">${s.value}</p><p class="text-xs text-gray-400 font-medium">${s.label}</p>${s.sub?`<p class="text-xs text-gray-400 mt-0.5">${s.sub}</p>`:''}</div>`).join('');
}

function renderCharts(){const ds=getActiveDataset();if(!ds)return;const ml=ds.metadata.monthAr||'—';
    const ps = document.getElementById('pie-section'); if(ps) ps.style.display=ds.summaryCols.length>0?'':'none';
    const c1 = document.getElementById('chart-month-label1'); if(c1) c1.textContent=ml;
    const c3 = document.getElementById('chart-month-label3'); if(c3) c3.textContent='القسم بأكمله';
    const c4 = document.getElementById('chart-month-label4'); if(c4) c4.textContent=ml;
    updateBarChart();updateLineChart();updatePieChart();updateHBarChart();updateStudentSelect();
}

function updateBarChart(){const ds=getActiveDataset();if(!ds)return;const sorted=[...ds.students].sort((a,b)=>{const va=state.barChartMode==='total'?calcTotalAbsences(a):getUnjustified(a);const vb=state.barChartMode==='total'?calcTotalAbsences(b):getUnjustified(b);return vb-va}).filter(st=>(state.barChartMode==='total'?calcTotalAbsences(st):getUnjustified(st))>0);if(barChart)barChart.destroy();
    const canvas = document.getElementById('barChart') as HTMLCanvasElement;
    if(!canvas) return;
    barChart=new (window as any).Chart(canvas,{type:'bar',data:{labels:sorted.map(st=>st.family),datasets:[{label:state.barChartMode==='total'?'المجموع':'غير مبرر',data:sorted.map(st=>state.barChartMode==='total'?calcTotalAbsences(st):getUnjustified(st)),backgroundColor:sorted.map((_,i)=>{const r=i/Math.max(sorted.length-1,1);return r<.5?CHART_COLORS[Math.floor(r*4)]:'#DC2626'}),borderRadius:6,maxBarThickness:32}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{rtl:true,textDirection:'rtl',titleFont:{family:'Cairo'},bodyFont:{family:'Cairo'}}},scales:{x:{ticks:{font:{family:'Cairo',size:10},maxRotation:45,minRotation:30},grid:{display:false}},y:{beginAtZero:true,ticks:{font:{family:'Cairo',size:11},stepSize:1},grid:{color:'#f0f1f4'}}}}})
}

function updateStudentSelect(){const sel=document.getElementById('line-chart-student') as HTMLSelectElement;if(!sel)return;const cv=sel.value;const ds=getActiveDataset();if(!ds)return;sel.innerHTML='<option value="all">جميع التلاميذ</option>';ds.students.forEach((st: any)=>{const o=document.createElement('option');o.value=st.id;o.textContent=studentFullName(st);sel.appendChild(o)});if(cv&&ds.students.some((st: any)=>st.id===cv))sel.value=cv}

function updateLineChart(){const ds=getActiveDataset();if(!ds)return;const cds=getActiveClassDatasets();const labels=cds.map((d: any)=>d.metadata.monthAr);const sel = document.getElementById('line-chart-student') as HTMLSelectElement; const sv=sel?sel.value:'all';let datasets: any[]=[];if(sv==='all'){datasets=[{label:'مجموع غيابات القسم',data:cds.map((d: any)=>d.students.reduce((s: number,st: any)=>s+calcTotalAbsences(st),0)),borderColor:'#059669',backgroundColor:'rgba(5,150,105,.1)',fill:true,tension:.4,pointRadius:5,pointHoverRadius:8,borderWidth:3}]}else{const tids=[sv,...[...ds.students].sort((a,b)=>calcTotalAbsences(b)-calcTotalAbsences(a)).filter(st=>st.id!==sv).slice(0,3).map(st=>st.id)];tids.forEach((sid,idx)=>{const f=ds.students.find((st: any)=>st.id===sid);if(!f)return;datasets.push({label:studentFullName(f),data:cds.map((d: any)=>{const s=d.students.find((x: any)=>x.id===sid);return s?calcTotalAbsences(s):0}),borderColor:CHART_COLORS[idx%CHART_COLORS.length],backgroundColor:'transparent',tension:.4,pointRadius:4,pointHoverRadius:7,borderWidth:2.5,borderDash:idx===0?[]:[5,3]})})}if(lineChart)lineChart.destroy();
    const canvas = document.getElementById('lineChart') as HTMLCanvasElement;
    if(!canvas) return;
    lineChart=new (window as any).Chart(canvas,{type:'line',data:{labels,datasets},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom',rtl:true,labels:{font:{family:'Cairo',size:11},usePointStyle:true,padding:16}},tooltip:{rtl:true,textDirection:'rtl',titleFont:{family:'Cairo'},bodyFont:{family:'Cairo'},mode:'index',intersect:false}},scales:{x:{ticks:{font:{family:'Cairo',size:11}},grid:{color:'#f0f1f4'}},y:{beginAtZero:true,ticks:{font:{family:'Cairo',size:11},stepSize:1},grid:{color:'#f0f1f4'}}},interaction:{mode:'nearest',axis:'x',intersect:false}}})
}

function updatePieChart(){
    const datasets=getActiveClassDatasets();
    if(!datasets||datasets.length===0)return;
    let tu=0,tj=0;
    datasets.forEach((ds: any) => {
        ds.students.forEach((st: any)=>{
            tu+=getUnjustified(st);
            tj+=getJustified(st);
        });
    });
    if(pieChart)pieChart.destroy();
    const canvas = document.getElementById('pieChart') as HTMLCanvasElement;
    if(!canvas) return;
    pieChart=new (window as any).Chart(canvas,{type:'doughnut',data:{labels:['غياب غير مبرر (أيام)','غياب مبرر (أيام)'],datasets:[{data:[tu,tj],backgroundColor:['#DC2626','#059669'],borderWidth:0,hoverOffset:8}]},options:{responsive:true,maintainAspectRatio:false,cutout:'60%',plugins:{legend:{position:'bottom',rtl:true,labels:{font:{family:'Cairo',size:12},usePointStyle:true,padding:20}},tooltip:{rtl:true,titleFont:{family:'Cairo'},bodyFont:{family:'Cairo'}}}}})
}

function updateHBarChart(){const ds=getActiveDataset();if(!ds)return;const sorted=[...ds.students].sort((a,b)=>calcTotalAbsences(b)-calcTotalAbsences(a)).slice(0,10);if(hBarChart)hBarChart.destroy();
    const canvas = document.getElementById('hBarChart') as HTMLCanvasElement;
    if(!canvas) return;
    hBarChart=new (window as any).Chart(canvas,{type:'bar',data:{labels:sorted.map(st=>studentFullName(st)),datasets:[{label:'المجموع',data:sorted.map(st=>calcTotalAbsences(st)),backgroundColor:sorted.map((_,i)=>i===0?'#DC2626':i<3?'#F59E0B':'#059669'),borderRadius:4,maxBarThickness:20}]},options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{rtl:true,titleFont:{family:'Cairo'},bodyFont:{family:'Cairo'}}},scales:{x:{beginAtZero:true,ticks:{font:{family:'Cairo',size:10},stepSize:1},grid:{color:'#f0f1f4'}},y:{ticks:{font:{family:'Cairo',size:10}},grid:{display:false}}}}})
}

function getStudentPhone(st: any): string {
    if (!st) return '';
    
    const id = st.id ? String(st.id).trim().toLowerCase() : '';
    const family = st.family ? String(st.family).trim().toLowerCase() : '';
    const name = st.name ? String(st.name).trim().toLowerCase() : '';
    
    const f1 = `${family} ${name}`.replace(/\s+/g, ' ').trim();
    const f2 = `${name} ${family}`.replace(/\s+/g, ' ').trim();
    
    // 1. Exact match on ID
    if (id && state.guardians[id]) return state.guardians[id];
    if (id && state.guardians[st.id]) return state.guardians[st.id]; // Fallback to original case
    
    // 2. Exact match on Name combinations
    if (state.guardians[f1]) return state.guardians[f1];
    if (state.guardians[f2]) return state.guardians[f2];
    
    const origF1 = `${st.family} ${st.name}`.replace(/\s+/g, ' ').trim();
    const origF2 = `${st.name} ${st.family}`.replace(/\s+/g, ' ').trim();
    if (state.guardians[origF1]) return state.guardians[origF1];
    if (state.guardians[origF2]) return state.guardians[origF2];
    
    // 3. Case-insensitive search across all keys
    for (const key in state.guardians) {
        const k = key.trim().toLowerCase();
        if (!k) continue;
        
        if (id && k === id) return state.guardians[key];
        if (k === f1 || k === f2) return state.guardians[key];
        
        // Sometimes names in Excel have extra spaces between words
        const normalizedK = k.replace(/\s+/g, ' ');
        if (normalizedK === f1 || normalizedK === f2) return state.guardians[key];
    }
    
    return '';
}

(window as any).updateStudentPhone = (studentId: string, phone: string) => {
    state.guardians[studentId.toLowerCase()] = phone;
    saveGuardiansToFirebase();
    showToast('تم حفظ رقم الهاتف', 'success');
    renderTable();
};

(window as any).enablePhoneEdit = (studentId: string, currentPhone: string) => {
    const td = document.getElementById(`phone-cell-${studentId}`);
    if (td) {
        td.innerHTML = `<input type="text"
            class="w-full px-3 py-2 border border-gray-200 rounded-lg focus:outline-none focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500 text-sm font-mono text-right"
            placeholder="أدخل رقم الهاتف..."
            value="${currentPhone}"
            onchange="updateStudentPhone('${studentId}', this.value)"
            onblur="renderGuardiansTable()" />`;
        const input = td.querySelector('input');
        if (input) input.focus();
    }
};

function getStudentGuardianInfo(st: any) {
    if (!st) return {};
    const id = st.id ? String(st.id).trim().toLowerCase() : '';
    const family = st.family ? String(st.family).trim().toLowerCase() : '';
    const name = st.name ? String(st.name).trim().toLowerCase() : '';
    const f1 = `${family} ${name}`.replace(/\s+/g, ' ').trim();
    const f2 = `${name} ${family}`.replace(/\s+/g, ' ').trim();
    
    return state.guardianDetails[id] || state.guardianDetails[f1] || state.guardianDetails[f2] || {};
}

function renderGuardiansTable() {
    const ds = getActiveDataset();
    const gtc = document.getElementById('guardians-table-card');
    const gtb = document.getElementById('guardians-table-body');
    
    if (!ds || !gtc || !gtb) return;
    
    let withPhone = 0;
    let withoutPhone = 0;
    ds.students.forEach((st: any) => {
        if (getStudentPhone(st)) withPhone++;
        else withoutPhone++;
    });
    
    const analysisEl = document.getElementById('guardians-analysis');
    if (analysisEl) {
        const total = ds.students.length;
        const percentage = total > 0 ? Math.round((withPhone / total) * 100) : 0;
        analysisEl.innerHTML = `
            <div class="flex items-center gap-2">
                <span class="text-emerald-600 bg-emerald-50 px-2 py-1 rounded-md"><i class="fa-solid fa-check mr-1"></i> ${withPhone} متوفر</span>
                <span class="text-rose-600 bg-rose-50 px-2 py-1 rounded-md"><i class="fa-solid fa-xmark mr-1"></i> ${withoutPhone} غير متوفر</span>
                <span class="text-blue-600 bg-blue-50 px-2 py-1 rounded-md ml-2">${percentage}% إنجاز</span>
            </div>
        `;
    }

    let students = [...ds.students];
    if (state.searchQuery) {
        const q = state.searchQuery.toLowerCase();
        students = students.filter(st => st.family.toLowerCase().includes(q) || st.name.toLowerCase().includes(q) || st.id.toLowerCase().includes(q));
    }
    
    let b = '';
    
    students.forEach((st, idx) => {
        const phone = getStudentPhone(st);
        const details = getStudentGuardianInfo(st);
        
        let detailsHtml = '';
        if (details.father) detailsHtml += `<div class="text-[10px] text-gray-400">الأب: ${details.father}</div>`;
        if (details.mother) detailsHtml += `<div class="text-[10px] text-gray-400">الأم: ${details.mother}</div>`;
        if (details.guardian) detailsHtml += `<div class="text-[10px] text-gray-400">الولي: ${details.guardian}</div>`;
        
        let phoneHtml = '';
        if (phone) {
            phoneHtml = `<div class="flex items-center justify-end gap-3">
                <span class="text-gray-700 font-mono text-sm">${phone}</span>
                <button onclick="enablePhoneEdit('${st.id}', '${phone}')" class="text-gray-400 hover:text-emerald-600 transition-colors" title="تعديل الرقم"><i class="fa-solid fa-pen text-xs"></i></button>
            </div>`;
        } else {
            phoneHtml = `<input type="text"
                   class="w-full px-3 py-2 border border-gray-200 rounded-lg focus:outline-none focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500 text-sm font-mono text-right"
                   placeholder="أدخل رقم الهاتف..."
                   onchange="updateStudentPhone('${st.id}', this.value)" />`;
        }
        
        b += `<tr class="fade-up" style="animation-delay:${Math.min(idx*20,400)}ms">
            <td class="text-right">
                <div class="font-bold text-gray-800">${st.family} ${st.name}</div>
                ${detailsHtml}
            </td>
            <td class="text-gray-500 text-sm">${st.id}</td>
            <td dir="ltr" id="phone-cell-${st.id}">
                ${phoneHtml}
            </td>
        </tr>`;
    });
    
    gtb.innerHTML = b;
    gtc.style.display = students.length > 0 ? '' : 'none';
}

function renderTable(){
    const ds=getActiveDataset();if(!ds)return;let students=[...ds.students];
    if(state.searchQuery){const q=state.searchQuery.toLowerCase();students=students.filter(st=>st.family.toLowerCase().includes(q)||st.name.toLowerCase().includes(q)||st.id.toLowerCase().includes(q))}
    if(state.sortMode==='absent')students.sort((a,b)=>calcTotalAbsences(b)-calcTotalAbsences(a));
    const nr = document.getElementById('no-results'); if(nr) nr.style.display=students.length===0?'':'none';
    const tc = document.getElementById('table-card'); if(tc) tc.style.display=students.length===0?'none':'';
    const sd=getSchoolDays(ds.students);const isD=state.viewMode==='detailed';
    let h=`<tr><th class="sticky-col right-0 text-right" style="min-width:45px">#</th><th class="sticky-col text-right" style="min-width:110px;right:45px">النسب</th><th class="sticky-col text-right" style="min-width:130px;right:155px">الإسم</th><th style="min-width:70px">المجموع</th>`;
    if(ds.summaryCols.length>0){['يوم غ.مبرر','يوم مبرر','ساعة غ.مبرر','ساعة مبرر'].forEach((l,i)=>{h+=`<th style="min-width:65px">${ds.summaryCols[i]?l:''}</th>`})}
    if(isD){for(let d=1;d<=31;d++){h+=`<th style="min-width:38px" class="${sd.has(d)?'':'opacity-40'}">${d}</th>`}}
    h+=`<th class="no-print" style="min-width:80px"></th></tr>`;
    const th = document.getElementById('table-head'); if(th) th.innerHTML=h;
    let b='';
    students.forEach((st,idx)=>{
        const ta=calcTotalAbsences(st),ad=calcAbsentDays(st),ha=ta>0;
        b+=`<tr class="fade-up" style="animation-delay:${Math.min(idx*20,400)}ms"><td class="sticky-col right-0 text-right font-semibold text-gray-500 text-xs" style="background:var(--card)">${st.rank}</td><td class="sticky-col text-right font-bold" style="background:var(--card);right:45px">${st.family}</td><td class="sticky-col text-right" style="background:var(--card);right:155px">${st.name}</td><td><span class="inline-flex items-center justify-center w-9 h-7 rounded-lg text-sm font-bold ${ha?'bg-red-50 text-red-600':'bg-emerald-50 text-emerald-600'}">${st.totalRaw!==''?st.totalRaw:ad}</span></td>`;
        ds.summaryCols.forEach((_: any,i: number)=>{const v=st.summaries[i]||'0';b+=`<td><span class="${parseInt(v)>0?'text-red-600 font-bold':'text-gray-400'}">${v}</span></td>`});
        if(isD){for(let d=1;d<=31;d++){const info=parseDayValue(st.days[d]);let cls='day-none';if(info.type==='present')cls='day-present';else if(info.type==='absent')cls=info.val>=2?'day-absent-high':'day-absent';else if(info.type==='special')cls='day-special';const disp=info.type==='none'?'—':(info.type==='present'?'0':(info.type==='special'?'X':info.val));b+=`<td><span class="day-cell ${cls}">${disp}</span></td>`}}
        b+=`<td class="no-print"><div class="flex items-center justify-center gap-1">`;
        
        // Add WhatsApp button if absences > 10
        if (ta > 10) {
            const phone = getStudentPhone(st);
            b+=`<button onclick="sendWhatsAppMessage('${st.id}', '${studentFullName(st).replace(/'/g, "\\'")}', ${ta}, '${phone}')" class="btn-ghost rounded-lg text-green-600 hover:bg-green-50" title="مراسلة عبر واتساب"><i class="fa-brands fa-whatsapp text-xs"></i></button>`;
        }
        
        b+=`<button onclick='openAbsenceSheet(${JSON.stringify(st.id).replace(/'/g,"\\'")})' class="btn-ghost rounded-lg text-amber-600 hover:bg-amber-50" title="ورقة الغياب"><i class="fa-solid fa-file-lines text-xs"></i></button><button onclick='openDetail(${JSON.stringify(st.id).replace(/'/g,"\\'")})' class="btn-ghost rounded-lg text-emerald-600 hover:bg-emerald-50" title="تفاصيل"><i class="fa-solid fa-arrow-left text-xs"></i></button></div></td></tr>`;
    });
    const tb = document.getElementById('table-body'); if(tb) tb.innerHTML=b;
    const tf = document.getElementById('table-footer'); if(tf) tf.innerHTML=`<span>عرض <strong>${students.length}</strong> من أصل <strong>${ds.students.length}</strong> تلميذ</span><span class="text-xs text-gray-400">${isD?'عرض تفصيلي — 31 يوم':'عرض مختصر'}</span>`;

    renderGuardiansTable();
    renderLatenessMainTable();
}

function renderLatenessMainTable() {
    const ltc = document.getElementById('lateness-main-table-card');
    const ltb = document.getElementById('lateness-main-table-body');
    
    if (!ltc || !ltb) return;
    
    if (!state.activeClass) {
        ltc.style.display = 'none';
        return;
    }
    
    let classRecords = state.latenessRecords.filter((r: any) => r.className === state.activeClass).sort((a: any, b: any) => new Date(b.date).getTime() - new Date(a.date).getTime());
    
    // Calculate lateness counts per student BEFORE filtering so the count is accurate for the student
    const latenessCounts: Record<string, number> = {};
    classRecords.forEach((r: any) => {
        latenessCounts[r.studentId] = (latenessCounts[r.studentId] || 0) + 1;
    });

    if (state.searchQuery) {
        const q = state.searchQuery.toLowerCase();
        classRecords = classRecords.filter((r: any) => r.studentName.toLowerCase().includes(q) || r.studentId.toLowerCase().includes(q));
    }
    
    if (classRecords.length === 0) {
        ltc.style.display = 'none';
        return;
    }
    
    let b = '';
    
    classRecords.forEach((r: any, idx: number) => {
        const count = latenessCounts[r.studentId];
        let countHtml = `<span class="bg-gray-100 text-gray-600 px-2 py-1 rounded-md text-xs font-bold">${count} تأخرات</span>`;
        if (count > 2) {
            countHtml = `<span class="bg-red-100 text-red-600 px-2 py-1 rounded-md text-xs font-bold">${count} تأخرات</span>`;
        } else if (count === 2) {
            countHtml = `<span class="bg-orange-100 text-orange-600 px-2 py-1 rounded-md text-xs font-bold">${count} تأخرات</span>`;
        } else if (count === 1) {
            countHtml = `<span class="bg-emerald-100 text-emerald-600 px-2 py-1 rounded-md text-xs font-bold">تأخر واحد</span>`;
        }
        
        b += `<tr class="fade-up" style="animation-delay:${Math.min(idx*20,400)}ms">
            <td class="text-right font-bold">${r.studentName}</td>
            <td class="text-gray-600 text-sm">${r.date}</td>
            <td><span class="bg-rose-100 text-rose-700 px-2 py-1 rounded-md font-bold text-xs">${r.minutes} دقيقة</span></td>
            <td class="text-gray-600 text-sm">${r.subject}</td>
            <td class="text-center">${countHtml}</td>
            <td class="no-print text-center">
                <button onclick="deleteLatenessRecord('${r.id}')" class="text-red-500 hover:text-red-700 w-8 h-8 rounded-full hover:bg-red-50 transition-colors"><i class="fa-solid fa-trash"></i></button>
            </td>
        </tr>`;
    });
    
    ltb.innerHTML = b;
    ltc.style.display = '';
}

function buildCombinedAbsenceSheet(classDatasets: any[], sid: string, sheetId: string) {
    const firstDs = classDatasets[0];
    const m = firstDs.metadata;
    const firstSt = firstDs.students.find((s: any) => s.id === sid);
    if (!firstSt) return '';

    const hasS = classDatasets.some(ds => ds.summaryCols && ds.summaryCols.length > 0);

    let s = `<div class="absence-sheet" id="${sheetId}">
        <div class="sheet-hdr">
            <h2>المملكة المغربية</h2>
            <h3>وزارة التربية الوطنية والتعليم الأولي والرياضة</h3>
            <h3>الأكاديمية الجهوية للتربية والتكوين — ${m.academy||'—'}</h3>
        </div>
        <div class="info-grid">
            <div class="info-cell"><span class="lbl">المؤسسة:</span><span class="val">${m.institution||'—'}</span></div>
            <div class="info-cell"><span class="lbl">السنة الدراسية:</span><span class="val">${m.year||'—'}</span></div>
            <div class="info-cell"><span class="lbl">المستوى:</span><span class="val">${m.level||'—'}</span></div>
            <div class="info-cell"><span class="lbl">القسم:</span><span class="val">${m.class||'—'}</span></div>
            <div class="info-cell"><span class="lbl">رقم التلميذ:</span><span class="val" style="font-weight:700">${firstSt.id}</span></div>
            <div class="info-cell"><span class="lbl">الإسم الكامل:</span><span class="val" style="font-weight:900">${firstSt.family} ${firstSt.name}</span></div>
        </div>
        <table class="sheet-tbl">
            <thead><tr>
                <th class="col-month" style="width:70px">الشهر</th>`;

    for (let d = 1; d <= 31; d++) {
        s += `<th class="col-day">${d}</th>`;
    }
    s += `<th class="col-sum">المجموع</th>`;
    if (hasS) {
        s += `<th class="col-s1">غير مبرر<br>يوم</th><th class="col-s2">مبرر<br>يوم</th><th class="col-s3">غير مبرر<br>ساعة</th><th class="col-s4">مبرر<br>ساعة</th>`;
    }
    s += `</tr>`;
    if (hasS) {
        s += `<tr class="sub-row"><th></th>`;
        for (let d = 1; d <= 31; d++) { s += `<th class="col-day">X/½</th>`; }
        s += `<th></th><th>X</th><th>X</th><th>X</th><th>X</th></tr>`;
    }
    s += `</thead><tbody>`;

    let totalAbsencesAllMonths = 0;
    let totalSummariesAllMonths = [0, 0, 0, 0];

    classDatasets.forEach((ds: any) => {
        const st = ds.students.find((x: any) => x.id === sid);
        if (!st) return;

        const sd = getSchoolDays(ds.students);

        s += `<tr>
            <td style="font-weight:800">${ds.metadata.monthAr || ds.metadata.month}</td>`;

        for (let d = 1; d <= 31; d++) {
            const info = parseDayValue(st.days[d]);
            let display = '', cls = '';
            if (!sd.has(d)) { cls = 'off-day'; }
            else if (info.type === 'none') { cls = ''; }
            else if (info.type === 'present') { display = '0'; cls = 'zero-v'; }
            else if (info.type === 'absent') { display = String(info.val); cls = 'abs-v'; }
            else if (info.type === 'special') { display = 'X'; cls = 'spec-v'; }
            s += `<td class="col-day ${cls}">${display}</td>`;
        }

        const total = calcTotalAbsences(st);
        totalAbsencesAllMonths += total;
        s += `<td class="col-sum" style="font-size:8pt;color:${total > 0 ? '#dc2626' : '#059669'}">${st.totalRaw !== '' ? st.totalRaw : total}</td>`;

        if (hasS) {
            for (let i = 0; i < 4; i++) {
                const v = st.summaries[i] || '0';
                const nv = parseInt(v) || 0;
                totalSummariesAllMonths[i] += nv;
                const bg = i % 2 === 0 ? 'col-s1' : 'col-s2';
                const clr = nv > 0 ? 'abs-v' : 'zero-v';
                s += `<td class="${bg} ${clr}">${v}</td>`;
            }
        }
        s += `</tr>`;
    });

    s += `<tr style="background-color: #f8fafc; font-weight: bold;">
        <td style="text-align: center;">المجموع السنوي</td>`;
    for (let d = 1; d <= 31; d++) {
        s += `<td></td>`;
    }
    s += `<td class="col-sum" style="color:${totalAbsencesAllMonths > 0 ? '#dc2626' : '#059669'}">${totalAbsencesAllMonths}</td>`;
    
    if (hasS) {
        for (let i = 0; i < 4; i++) {
            const nv = totalSummariesAllMonths[i];
            const bg = i % 2 === 0 ? 'col-s1' : 'col-s2';
            const clr = nv > 0 ? 'abs-v' : 'zero-v';
            s += `<td class="${bg} ${clr}">${nv}</td>`;
        }
    }
    s += `</tr>`;

    s += `</tbody></table>
        <div class="ft-section">
            <div class="ft-row"><div class="ft-cell" style="grid-column:1/-1"><span class="fl">ملاحظات:</span><div class="fl-line"></div></div></div>
            <div class="ft-row">
                <div class="ft-cell"><span class="fl">توقيع الأستاذ(ة):</span><div class="fl-line"></div></div>
                <div class="ft-cell"><span class="fl">توقيع الحارس العام:</span><div class="fl-line"></div></div>
                <div class="ft-cell"><span class="fl">توقيع المدير(ة):</span><div class="fl-line"></div></div>
            </div>
        </div>
    </div>`;
    return s;
}

function buildPrintArea(ds: any, single: any) {
    const m = ds.metadata, sd = getSchoolDays(ds.students), students = single ? [single] : ds.students;
    let h = `<div style="font-family:Arial,sans-serif;direction:rtl"><div style="text-align:center;margin-bottom:16px"><h2 style="font-size:18px;font-weight:800;margin:0 0 4px">ملخص الغياب الشهري للمؤسسة</h2></div><table style="width:100%;margin-bottom:12px;border:none;font-size:12px"><tr><td style="border:none;padding:2px 8px;text-align:right"><strong>الأكاديمية:</strong> ${m.academy || '—'}</td><td style="border:none;padding:2px 8px;text-align:right"><strong>المؤسسة:</strong> ${m.institution || '—'}</td><td style="border:none;padding:2px 8px;text-align:right"><strong>السنة الدراسية:</strong> ${m.year || '—'}</td></tr><tr><td style="border:none;padding:2px 8px;text-align:right"><strong>المستوى:</strong> ${m.level || '—'}</td><td style="border:none;padding:2px 8px;text-align:right"><strong>القسم:</strong> ${m.class || '—'}</td><td style="border:none;padding:2px 8px;text-align:right"><strong>الشهر:</strong> ${m.monthAr || '—'}</td></tr></table><table style="width:100%;border-collapse:collapse;font-size:11px"><thead><tr><th style="border:1px solid #333;padding:4px 6px;background:#e8e8e8">#</th><th style="border:1px solid #333;padding:4px 6px;background:#e8e8e8">رقم التلميذ</th><th style="border:1px solid #333;padding:4px 6px;background:#e8e8e8">النسب</th><th style="border:1px solid #333;padding:4px 6px;background:#e8e8e8">الإسم</th>`;
    for (let d = 1; d <= 31; d++) { h += `<th style="border:1px solid #333;padding:4px 6px;background:${!sd.has(d) ? '#e8e8e8;color:#aaa' : ''}">${d}</th>`; }
    h += `<th style="border:1px solid #333;padding:4px 6px;background:#e8e8e8">المجموع</th>`;
    ds.summaryCols.forEach((_: any, i: number) => { h += `<th style="border:1px solid #333;padding:4px 6px;background:#e8e8e8">${['يوم غ.مبرر', 'يوم مبرر', 'ساعة غ.مبرر', 'ساعة مبرر'][i] || ''}</th>`; });
    h += `</tr></thead><tbody>`;
    students.forEach((st: any) => {
        h += `<tr><td style="border:1px solid #333;padding:4px 6px">${st.rank}</td><td style="border:1px solid #333;padding:4px 6px">${st.id}</td><td style="border:1px solid #333;padding:4px 6px">${st.family}</td><td style="border:1px solid #333;padding:4px 6px">${st.name}</td>`;
        for (let d = 1; d <= 31; d++) { const info = parseDayValue(st.days[d]); const disp = info.type === 'none' ? '' : (info.type === 'present' ? '0' : (info.type === 'special' ? 'X' : info.val)); h += `<td style="border:1px solid #333;padding:4px 6px;${info.type === 'absent' ? 'background:#fee2e2;font-weight:700;color:#dc2626' : ''}">${disp}</td>`; }
        h += `<td style="border:1px solid #333;padding:4px 6px;font-weight:700">${st.totalRaw || calcAbsentDays(st)}</td>`;
        ds.summaryCols.forEach((_: any, i: number) => { h += `<td style="border:1px solid #333;padding:4px 6px">${st.summaries[i] || '0'}</td>`; });
        h += `</tr>`;
    });
    h += `</tbody></table><p style="margin-top:12px;font-size:10px;color:#888;text-align:center">نظام إدارة الغيابات — ${new Date().toLocaleDateString('ar-MA')}</p></div>`;
    const pa = document.getElementById('print-area'); if(pa) { pa.innerHTML = h; pa.style.display = 'block'; }
}

// Global functions for HTML
(window as any).setPrintOrientation = (orientation: 'portrait' | 'landscape') => {
    let styleEl = document.getElementById('print-orientation-style');
    if (!styleEl) {
        styleEl = document.createElement('style');
        styleEl.id = 'print-orientation-style';
        document.head.appendChild(styleEl);
    }
    styleEl.innerHTML = `@page { size: ${orientation}; }`;
};

(window as any).closeDetailPanel = () => {
    const dp = document.getElementById('detail-panel'); if(dp) dp.classList.remove('open');
    const po = document.getElementById('panel-overlay'); if(po) po.classList.remove('open');
};

(window as any).closeSheetModal = () => {
    const sm = document.getElementById('sheet-modal'); if(sm) sm.classList.remove('open');
};

(window as any).printAbsenceSheet = () => {
    const sheets = document.querySelectorAll('[id^="absence-sheet-el"]');
    if (!sheets.length) return;
    const pa = document.getElementById('sheet-print-area');
    if(pa) {
        pa.innerHTML = Array.from(sheets).map(el => el.outerHTML).join('');
        pa.style.display = 'block';
        (window as any).setPrintOrientation('landscape');
        setTimeout(() => window.print(), 250);
    }
};

let currentCommitmentStudents: any[] = [];

(window as any).closeCommitmentsModal = () => {
    const m = document.getElementById('commitments-modal');
    if (m) m.classList.remove('open');
};

(window as any).openParentsVisitsModal = () => {
    const classDatasets = getActiveClassDatasets();
    if (!classDatasets.length) {
        showToast('لا توجد بيانات', 'error');
        return;
    }
    
    const firstDs = classDatasets[0];
    const m = firstDs.metadata;
    
    let html = `
    <div id="parents-visits-print-content" class="bg-white p-8" style="direction: rtl;">
        <div class="text-center mb-6">
            <h2 class="text-xl font-bold mb-2">سجل زيارات الآباء وأولياء الأمور</h2>
            <div class="flex justify-center gap-6 text-sm font-semibold text-gray-700">
                <span>المؤسسة: ${m.institution || '—'}</span>
                <span>المستوى: ${m.level || '—'}</span>
                <span>القسم: ${m.className || '—'}</span>
                <span>الموسم الدراسي: ${m.year || '—'}</span>
            </div>
        </div>
        
        <table class="w-full border-collapse border border-gray-800 text-sm text-center">
            <thead>
                <tr class="bg-gray-100">
                    <th class="border border-gray-800 p-2 w-10">الرقم</th>
                    <th class="border border-gray-800 p-2 w-48">اسم التلميذ(ة)</th>
                    <th class="border border-gray-800 p-2 w-28">الرقم الوطني</th>
                    <th class="border border-gray-800 p-2 w-48">اسم الولي</th>
                    <th class="border border-gray-800 p-2 w-40">سبب الزيارة</th>
                    <th class="border border-gray-800 p-2 w-28">تاريخ الزيارة</th>
                    <th class="border border-gray-800 p-2 w-24">التوقيع</th>
                    <th class="border border-gray-800 p-2">ملاحظات</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    // Aggregate absences across all months
    const studentAbsences = new Map<string, { student: any, totalAbsences: number }>();
    
    classDatasets.forEach((ds: any) => {
        ds.students.forEach((st: any) => {
            const absences = calcTotalAbsences(st);
            if (studentAbsences.has(st.id)) {
                studentAbsences.get(st.id)!.totalAbsences += absences;
            } else {
                studentAbsences.set(st.id, { student: st, totalAbsences: absences });
            }
        });
    });

    // Filter students with more than 10 hours of absence across all months
    const exceededStudents = Array.from(studentAbsences.values())
        .filter(item => item.totalAbsences > 10)
        .map(item => ({ ...item.student, aggregatedAbsences: item.totalAbsences }));

    const totalRows = Math.max(20, exceededStudents.length);

    for (let i = 1; i <= totalRows; i++) {
        const st = exceededStudents[i - 1];
        if (st) {
            const gd = state.guardianDetails[st.id] || {};
            const guardianName = gd.guardian || gd.father || gd.mother || '';
            html += `
                <tr>
                    <td class="border border-gray-800 p-2 h-10">${i}</td>
                    <td class="border border-gray-800 p-2">${studentFullName(st)}</td>
                    <td class="border border-gray-800 p-2">${st.id || ''}</td>
                    <td class="border border-gray-800 p-2">${guardianName}</td>
                    <td class="border border-gray-800 p-2">تجاوز الغياب المسموح (${st.aggregatedAbsences} س)</td>
                    <td class="border border-gray-800 p-2"></td>
                    <td class="border border-gray-800 p-2"></td>
                    <td class="border border-gray-800 p-2"></td>
                </tr>
            `;
        } else {
            html += `
                <tr>
                    <td class="border border-gray-800 p-2 h-10">${i}</td>
                    <td class="border border-gray-800 p-2"></td>
                    <td class="border border-gray-800 p-2"></td>
                    <td class="border border-gray-800 p-2"></td>
                    <td class="border border-gray-800 p-2"></td>
                    <td class="border border-gray-800 p-2"></td>
                    <td class="border border-gray-800 p-2"></td>
                    <td class="border border-gray-800 p-2"></td>
                </tr>
            `;
        }
    }
    
    html += `
            </tbody>
        </table>
    </div>
    `;
    
    const container = document.getElementById('parents-visits-content');
    if (container) {
        container.innerHTML = html;
    }
    
    const modal = document.getElementById('parents-visits-modal');
    if (modal) modal.classList.add('open');
};

(window as any).closeParentsVisitsModal = () => {
    const modal = document.getElementById('parents-visits-modal');
    if (modal) modal.classList.remove('open');
};

(window as any).openLatenessModal = () => {
    if (!state.activeClass) {
        showToast('الرجاء اختيار قسم أولاً', 'error');
        return;
    }
    
    const select = document.getElementById('lateness-student-select');
    if (select) {
        const ds = getActiveDataset();
        if (ds) {
            select.innerHTML = ds.students.map((st: any) => `<option value="${st.id}">${st.family} ${st.name}</option>`).join('');
        }
    }
    
    const dateInput = document.getElementById('lateness-date') as HTMLInputElement;
    if (dateInput) {
        dateInput.value = new Date().toISOString().split('T')[0];
    }
    
    renderLatenessTable();
    
    const m = document.getElementById('lateness-modal');
    if (m) m.classList.add('open');
};

(window as any).closeLatenessModal = () => {
    const m = document.getElementById('lateness-modal');
    if (m) m.classList.remove('open');
};

(window as any).addLatenessRecord = () => {
    const studentId = (document.getElementById('lateness-student-select') as HTMLSelectElement)?.value;
    const minutes = (document.getElementById('lateness-minutes') as HTMLInputElement)?.value;
    const date = (document.getElementById('lateness-date') as HTMLInputElement)?.value;
    const subject = (document.getElementById('lateness-subject') as HTMLInputElement)?.value;
    
    if (!studentId || !minutes || !date || !subject) {
        showToast('الرجاء ملء جميع الحقول', 'error');
        return;
    }
    
    const ds = getActiveDataset();
    const student = ds?.students.find((s: any) => s.id === studentId);
    
    if (!student) return;
    
    const record = {
        id: Date.now().toString(),
        studentId,
        studentName: `${student.family} ${student.name}`,
        className: state.activeClass,
        minutes: parseInt(minutes),
        date,
        subject
    };
    
    state.latenessRecords.push(record);
    saveDataLocally();
    renderLatenessTable();
    
    (document.getElementById('lateness-minutes') as HTMLInputElement).value = '15';
    (document.getElementById('lateness-subject') as HTMLInputElement).value = '';
    
    showToast('تم تسجيل التأخر بنجاح', 'success');
};

(window as any).deleteLatenessRecord = (id: string) => {
    (window as any).showConfirmModal('هل أنت متأكد من حذف هذا السجل؟', () => {
        state.latenessRecords = state.latenessRecords.filter((r: any) => r.id !== id);
        saveDataLocally();
        renderLatenessTable();
        showToast('تم الحذف بنجاح', 'success');
    });
};

function renderLatenessTable() {
    renderLatenessMainTable();
    const container = document.getElementById('lateness-content');
    if (!container) return;
    
    const classRecords = state.latenessRecords.filter((r: any) => r.className === state.activeClass).sort((a: any, b: any) => new Date(b.date).getTime() - new Date(a.date).getTime());
    
    if (classRecords.length === 0) {
        container.innerHTML = '<div class="text-center text-gray-400 py-8">لا توجد تأخرات مسجلة لهذا القسم</div>';
        return;
    }
    
    // Calculate lateness counts per student
    const latenessCounts: Record<string, number> = {};
    classRecords.forEach((r: any) => {
        latenessCounts[r.studentId] = (latenessCounts[r.studentId] || 0) + 1;
    });
    
    let html = `
    <table class="w-full text-sm text-right" style="border-collapse: collapse;">
        <thead class="bg-gray-50 text-gray-600 font-bold border-b border-gray-200">
            <tr>
                <th class="py-3 px-4" style="border: 1px solid #e5e7eb;">التلميذ(ة)</th>
                <th class="py-3 px-4" style="border: 1px solid #e5e7eb;">التاريخ</th>
                <th class="py-3 px-4" style="border: 1px solid #e5e7eb;">المدة</th>
                <th class="py-3 px-4" style="border: 1px solid #e5e7eb;">المادة</th>
                <th class="py-3 px-4" style="border: 1px solid #e5e7eb;">عدد التأخرات</th>
                <th class="py-3 px-4 no-print text-center" style="border: 1px solid #e5e7eb;">إجراء</th>
            </tr>
        </thead>
        <tbody>
    `;
    
    classRecords.forEach((r: any) => {
        const count = latenessCounts[r.studentId];
        let countHtml = `<span class="bg-gray-100 text-gray-600 px-2 py-1 rounded-md text-xs font-bold">${count} تأخرات</span>`;
        if (count > 2) {
            countHtml = `<span class="bg-red-100 text-red-600 px-2 py-1 rounded-md text-xs font-bold">${count} تأخرات</span>`;
        } else if (count === 2) {
            countHtml = `<span class="bg-orange-100 text-orange-600 px-2 py-1 rounded-md text-xs font-bold">${count} تأخرات</span>`;
        } else if (count === 1) {
            countHtml = `<span class="bg-emerald-100 text-emerald-600 px-2 py-1 rounded-md text-xs font-bold">تأخر واحد</span>`;
        }
        
        html += `
            <tr class="border-b border-gray-100 hover:bg-gray-50">
                <td class="py-3 px-4 font-bold text-gray-800" style="border: 1px solid #e5e7eb;">${r.studentName}</td>
                <td class="py-3 px-4 text-gray-600" style="border: 1px solid #e5e7eb;">${r.date}</td>
                <td class="py-3 px-4" style="border: 1px solid #e5e7eb;"><span class="bg-rose-100 text-rose-700 px-2 py-1 rounded-md font-bold">${r.minutes} دقيقة</span></td>
                <td class="py-3 px-4 text-gray-600" style="border: 1px solid #e5e7eb;">${r.subject}</td>
                <td class="py-3 px-4 text-center" style="border: 1px solid #e5e7eb;">${countHtml}</td>
                <td class="py-3 px-4 no-print text-center" style="border: 1px solid #e5e7eb;">
                    <button onclick="deleteLatenessRecord('${r.id}')" class="text-red-500 hover:text-red-700 w-8 h-8 rounded-full hover:bg-red-50 transition-colors"><i class="fa-solid fa-trash"></i></button>
                </td>
            </tr>
        `;
    });
    
    html += `</tbody></table>`;
    container.innerHTML = html;
}

(window as any).executePrintLateness = () => {
    const content = document.getElementById('lateness-content');
    const pa = document.getElementById('sheet-print-area');
    if (content && pa) {
        pa.innerHTML = `
            <div style="padding: 20px; font-family: 'Cairo', sans-serif; direction: rtl;">
                <h2 style="text-align: center; margin-bottom: 10px; color: #111;">سجل تأخرات التلاميذ</h2>
                <h3 style="text-align: center; margin-bottom: 20px; color: #555;">القسم: ${state.activeClass}</h3>
                ${content.innerHTML}
            </div>
        `;
        pa.style.display = 'block';
        (window as any).setPrintOrientation('portrait');
        (window as any).closeLatenessModal();
        setTimeout(() => window.print(), 500);
    }
};

(window as any).executePrintParentsVisits = () => {
    const content = document.getElementById('parents-visits-print-content');
    if (content) {
        const pa = document.getElementById('print-area');
        if (!pa) return;
        pa.innerHTML = content.outerHTML;
        pa.style.display = 'block';
        (window as any).setPrintOrientation('landscape');
        (window as any).closeParentsVisitsModal();
        setTimeout(() => window.print(), 500);
    }
};

(window as any).downloadParentsVisitsPDF = async () => {
    const content = document.getElementById('parents-visits-print-content');
    if (!content) return;
    
    showToast('جارٍ إنشاء ملف PDF (قد يستغرق بعض الوقت)...', 'info', 10000);
    
    try {
        const opt = {
            margin: 10,
            filename: `سجل_زيارات_الآباء.pdf`,
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { scale: 2, useCORS: true },
            jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape' }
        };
        
        await (window as any).html2pdf().set(opt).from(content).save();
        showToast('تم تحميل الملف بنجاح', 'success');
    } catch (error) {
        console.error(error);
        showToast('حدث خطأ أثناء إنشاء ملف PDF', 'error');
    }
};

(window as any).printCommitments = () => {
    const classDatasets = getActiveClassDatasets();
    if (!classDatasets.length) {
        showToast('لا توجد بيانات', 'error');
        return;
    }
    
    const firstDs = classDatasets[0];
    
    // Aggregate absences across all months
    const studentAbsences = new Map<string, { student: any, totalAbsences: number }>();
    
    classDatasets.forEach((ds: any) => {
        ds.students.forEach((st: any) => {
            const absences = calcTotalAbsences(st);
            if (studentAbsences.has(st.id)) {
                studentAbsences.get(st.id)!.totalAbsences += absences;
            } else {
                studentAbsences.set(st.id, { student: st, totalAbsences: absences });
            }
        });
    });

    // Filter students with more than 10 hours of absence across all months
    currentCommitmentStudents = Array.from(studentAbsences.values())
        .filter(item => item.totalAbsences > 10)
        .map(item => ({ ...item.student, aggregatedAbsences: item.totalAbsences }));
    
    if (currentCommitmentStudents.length === 0) {
        showToast('لا يوجد تلاميذ تجاوزوا 10 ساعات من الغياب في هذا القسم', 'info');
        return;
    }

    const listContainer = document.getElementById('commitments-list');
    if (listContainer) {
        listContainer.innerHTML = currentCommitmentStudents.map((st: any, index: number) => `
            <div id="commitment-sheet-${index}" class="bg-white p-8 rounded-xl shadow-sm border border-gray-200 mb-6" style="font-family: 'Cairo', sans-serif; direction: rtl; line-height: 2;">
                <div style="text-align: center; margin-bottom: 40px;">
                    <h2 style="font-size: 24px; font-weight: bold; text-decoration: underline;">التزام ولي الأمر بخصوص المواظبة</h2>
                </div>
                
                <div style="font-size: 18px; margin-bottom: 30px;">
                    <p>أنا الموقع(ة) أسفله: ........................................................................</p>
                    <p>بصفتي ولي أمر التلميذ(ة): <strong>${studentFullName(st)}</strong></p>
                    <p>المسجل(ة) بالقسم: <strong>${firstDs.metadata.class}</strong></p>
                    <p>رقم التلميذ (مسار): <strong>${st.massar || '....................'}</strong></p>
                </div>

                <div style="font-size: 18px; margin-bottom: 40px; text-align: justify;">
                    <p>أقر بأني قد أُشعرت من طرف إدارة المؤسسة بتجاوز ابني/ابنتي لـ <strong class="text-red-600">10 ساعات</strong> من الغياب (مجموع الغيابات المسجلة: <strong class="text-red-600">${st.aggregatedAbsences}</strong> ساعة).</p>
                    <p>وبناءً عليه، <strong>ألتزم</strong> بالحرص على مواظبة ابني/ابنتي على الحضور في الأوقات المحددة للدراسة، وتبرير أي غياب مستقبلي بوثائق رسمية (شهادة طبية، إلخ) في الآجال القانونية.</p>
                    <p>وفي حالة تمادي التلميذ(ة) في الغياب غير المبرر، أتحمل كامل المسؤولية المترتبة عن ذلك وفقاً لمقتضيات النظام الداخلي للمؤسسة.</p>
                </div>

                <div style="display: flex; justify-content: space-between; font-size: 18px; margin-top: 50px;">
                    <div>
                        <p>حرر بـ .................... في ....................</p>
                        <p style="margin-top: 20px; font-weight: bold;">توقيع ولي الأمر:</p>
                    </div>
                    <div>
                        <p style="font-weight: bold;">توقيع وإدارة المؤسسة:</p>
                    </div>
                </div>
            </div>
        `).join('');
    }

    const m = document.getElementById('commitments-modal');
    if (m) m.classList.add('open');
};

(window as any).executePrintCommitments = () => {
    const classDatasets = getActiveClassDatasets();
    if (!classDatasets.length || currentCommitmentStudents.length === 0) return;
    const firstDs = classDatasets[0];

    const pa = document.getElementById('sheet-print-area');
    if (!pa) return;

    let html = '';
    currentCommitmentStudents.forEach((st: any) => {
        html += `
            <div style="page-break-after: always; padding: 40px; font-family: 'Cairo', sans-serif; direction: rtl; line-height: 2;">
                <div style="text-align: center; margin-bottom: 40px;">
                    <h2 style="font-size: 24px; font-weight: bold; text-decoration: underline;">التزام ولي الأمر بخصوص المواظبة</h2>
                </div>
                
                <div style="font-size: 18px; margin-bottom: 30px;">
                    <p>أنا الموقع(ة) أسفله: ........................................................................</p>
                    <p>بصفتي ولي أمر التلميذ(ة): <strong>${studentFullName(st)}</strong></p>
                    <p>المسجل(ة) بالقسم: <strong>${firstDs.metadata.class}</strong></p>
                    <p>رقم التلميذ (مسار): <strong>${st.massar || '....................'}</strong></p>
                </div>

                <div style="font-size: 18px; margin-bottom: 40px; text-align: justify;">
                    <p>أقر بأني قد أُشعرت من طرف إدارة المؤسسة بتجاوز ابني/ابنتي لـ <strong>10 ساعات</strong> من الغياب (مجموع الغيابات المسجلة: <strong>${st.aggregatedAbsences}</strong> ساعة).</p>
                    <p>وبناءً عليه، <strong>ألتزم</strong> بالحرص على مواظبة ابني/ابنتي على الحضور في الأوقات المحددة للدراسة، وتبرير أي غياب مستقبلي بوثائق رسمية (شهادة طبية، إلخ) في الآجال القانونية.</p>
                    <p>وفي حالة تمادي التلميذ(ة) في الغياب غير المبرر، أتحمل كامل المسؤولية المترتبة عن ذلك وفقاً لمقتضيات النظام الداخلي للمؤسسة.</p>
                </div>

                <div style="display: flex; justify-content: space-between; font-size: 18px; margin-top: 50px;">
                    <div>
                        <p>حرر بـ .................... في ....................</p>
                        <p style="margin-top: 20px; font-weight: bold;">توقيع ولي الأمر:</p>
                    </div>
                    <div>
                        <p style="font-weight: bold;">توقيع وإدارة المؤسسة:</p>
                    </div>
                </div>
            </div>
        `;
    });

    pa.innerHTML = html;
    pa.style.display = 'block';
    (window as any).setPrintOrientation('portrait');
    (window as any).closeCommitmentsModal();
    setTimeout(() => window.print(), 500);
};

(window as any).closeWhatsAppModal = () => {
    const m = document.getElementById('whatsapp-modal');
    if (m) m.classList.remove('open');
};

(window as any).openWhatsAppModal = () => {
    const classDatasets = getActiveClassDatasets();
    if (!classDatasets.length) {
        showToast('لا توجد بيانات', 'error');
        return;
    }
    
    // Aggregate absences across all months
    const studentAbsences = new Map<string, { student: any, totalAbsences: number }>();
    
    classDatasets.forEach((ds: any) => {
        ds.students.forEach((st: any) => {
            const absences = calcTotalAbsences(st);
            if (studentAbsences.has(st.id)) {
                studentAbsences.get(st.id)!.totalAbsences += absences;
            } else {
                studentAbsences.set(st.id, { student: st, totalAbsences: absences });
            }
        });
    });

    // Filter students with more than 10 hours of absence across all months
    const studentsToNotify = Array.from(studentAbsences.values())
        .filter(item => item.totalAbsences > 10)
        .map(item => ({ ...item.student, aggregatedAbsences: item.totalAbsences }));
    
    if (studentsToNotify.length === 0) {
        showToast('لا يوجد تلاميذ تجاوزوا 10 ساعات من الغياب في هذا القسم', 'info');
        return;
    }

    const listContainer = document.getElementById('whatsapp-list');
    if (listContainer) {
        listContainer.innerHTML = studentsToNotify.map((st: any, index: number) => {
            const phone = getStudentPhone(st);
            return `
            <div class="bg-white p-4 rounded-xl shadow-sm border border-gray-200 flex flex-col gap-3">
                <div class="flex items-center justify-between">
                    <div class="flex items-center gap-3">
                        <div class="w-8 h-8 rounded-full bg-green-100 text-green-600 flex items-center justify-center font-bold text-sm">
                            ${index + 1}
                        </div>
                        <div>
                            <p class="font-bold text-gray-800 text-sm">${studentFullName(st)}</p>
                            <p class="text-xs text-gray-500">مجموع الغياب التراكمي: <span class="text-red-600 font-bold">${st.aggregatedAbsences} ساعة</span></p>
                        </div>
                    </div>
                </div>
                <div class="flex items-center gap-2 mt-2">
                    <div class="relative flex-1">
                        <i class="fa-solid fa-phone absolute right-3 top-1/2 -translate-y-1/2 text-gray-400 text-xs"></i>
                        <input type="tel" id="wa-phone-${st.id}" value="${phone}" placeholder="رقم هاتف ولي الأمر (مثال: 212600000000+)" class="w-full text-sm border border-gray-200 rounded-lg pr-8 pl-3 py-2 focus:outline-none focus:border-green-400" dir="ltr">
                    </div>
                    <button onclick="sendWhatsAppMessage('${st.id}', '${studentFullName(st).replace(/'/g, "\\'")}', ${st.aggregatedAbsences})" class="btn btn-primary bg-green-600 hover:bg-green-700 border-none text-xs py-2 px-4 whitespace-nowrap">
                        <i class="fa-regular fa-paper-plane ml-1"></i> إرسال
                    </button>
                </div>
            </div>
        `}).join('');
    }

    const m = document.getElementById('whatsapp-modal');
    if (m) m.classList.add('open');
};

(window as any).sendWhatsAppMessage = (studentId: string, studentName: string, absences: number, phoneOverride?: string) => {
    let phoneStr = phoneOverride;
    const phoneInput = document.getElementById(`wa-phone-${studentId}`) as HTMLInputElement;
    
    if (phoneInput) {
        phoneStr = phoneInput.value.trim();
    } else if (!phoneStr) {
        // Try to get from state if not provided
        for (const ds of state.datasets) {
            const st = ds.students.find((s: any) => s.id === studentId);
            if (st) {
                phoneStr = getStudentPhone(st);
                break;
            }
        }
    }
    
    if (!phoneStr) {
        const userInput = prompt(`لم يتم العثور على رقم هاتف للتلميذ ${studentName}.\nالمرجو إدخال رقم الهاتف:`);
        if (userInput && userInput.trim()) {
            phoneStr = userInput.trim();
        } else {
            return;
        }
    }

    // Save phone number to state if it's new or changed
    const idLower = studentId.toLowerCase();
    if (state.guardians[idLower] !== phoneStr) {
        state.guardians[idLower] = phoneStr;
        saveGuardiansToFirebase();
    }

    // Extract first valid phone number
    let phone = '';
    const phoneMatch = phoneStr.match(/(?:\+|00)?(?:212|0)[567]\d{8}/) || phoneStr.match(/\+?\d{9,15}/);
    if (phoneMatch) {
        phone = phoneMatch[0];
    } else {
        phone = phoneStr.replace(/[^0-9+]/g, '');
    }

    // Basic phone number cleaning
    phone = phone.replace(/[^0-9+]/g, '');
    if (phone.startsWith('00')) phone = '+' + phone.substring(2);
    if (phone.startsWith('0')) {
        phone = '212' + phone.substring(1); // Assuming Morocco country code by default if starts with 0
    }

    const message = `السلام عليكم،\nننهي إلى علمكم أن ابنكم/ابنتكم *${studentName}* قد تجاوز 10 ساعات من الغياب (مجموع الغيابات: ${absences} ساعة).\nالمرجو الحضور إلى المؤسسة لتسوية وضعيته.\nالإدارة.`;
    const encodedMessage = encodeURIComponent(message);
    
    const waUrl = `https://wa.me/${phone}?text=${encodedMessage}`;
    window.open(waUrl, '_blank');
};

(window as any).downloadCommitmentsPDF = async () => {
    if (currentCommitmentStudents.length === 0) return;

    showToast('جارٍ إنشاء ملف PDF (قد يستغرق بعض الوقت)...', 'info', 10000);
    
    try {
        const { jsPDF } = (window as any).jspdf;
        const pdf = new jsPDF('p', 'mm', 'a4');
        const imgW = 210;

        for (let i = 0; i < currentCommitmentStudents.length; i++) {
            if (i > 0) pdf.addPage();
            const sheet = document.getElementById(`commitment-sheet-${i}`);
            if (sheet) {
                const canvas = await (window as any).html2canvas(sheet, {
                    scale: 2,
                    useCORS: true,
                    logging: false,
                    backgroundColor: '#ffffff'
                });
                const imgData = canvas.toDataURL('image/jpeg', 1.0);
                const imgH = (canvas.height * imgW) / canvas.width;
                pdf.addImage(imgData, 'JPEG', 0, 0, imgW, imgH);
            }
        }

        pdf.save(`التزامات_الغياب_${getActiveDataset()?.metadata.class}.pdf`);
        showToast('تم تحميل ملف PDF بنجاح', 'success');
    } catch (err) {
        console.error(err);
        showToast('حدث خطأ أثناء إنشاء PDF', 'error');
    }
};

(window as any).exportAllStudentsSheetsExcel = () => {
    const classDatasets = getActiveClassDatasets();
    if (!classDatasets.length) return;

    const firstDs = classDatasets[0];
    const students = firstDs.students;
    const m = firstDs.metadata;

    if (!students || !students.length) {
        showToast('لا يوجد تلاميذ في هذا القسم', 'error');
        return;
    }

    showToast('جارٍ تجهيز الأوراق للتصدير...', 'info');

    try {
        const wb = (window as any).XLSX.utils.book_new();
        const hasS = classDatasets.some(ds => ds.summaryCols && ds.summaryCols.length > 0);

        students.forEach((st: any) => {
            const data = [];
            
            data.push(['المملكة المغربية']);
            data.push(['وزارة التربية الوطنية والتعليم الأولي والرياضة']);
            data.push([`الأكاديمية الجهوية للتربية والتكوين — ${m.academy || '—'}`]);
            data.push([]);
            data.push(['المؤسسة:', m.institution || '—', 'السنة الدراسية:', m.year || '—']);
            data.push(['المستوى:', m.level || '—', 'القسم:', m.class || '—']);
            data.push(['رقم التلميذ:', st.id, 'الإسم الكامل:', `${st.family} ${st.name}`]);
            data.push([]);

            const headers = ['الشهر'];
            for (let d = 1; d <= 31; d++) headers.push(d.toString());
            headers.push('المجموع');
            if (hasS) {
                headers.push('يوم غ.م');
                headers.push('يوم م');
                headers.push('ساعة غ.م');
                headers.push('ساعة م');
            }
            data.push(headers);

            classDatasets.forEach(ds => {
                const studentInMonth = ds.students.find((s: any) => s.id === st.id);
                if (!studentInMonth) return;

                const row = [ds.metadata.monthAr || ds.metadata.month];
                for (let d = 1; d <= 31; d++) {
                    const info = parseDayValue(studentInMonth.days[d]);
                    const disp = info.type === 'none' ? '' : (info.type === 'present' ? '0' : (info.type === 'special' ? 'X' : info.val));
                    row.push(disp);
                }
                row.push(calcTotalAbsences(studentInMonth));
                if (hasS) {
                    row.push(studentInMonth.summaries[0] || '0');
                    row.push(studentInMonth.summaries[1] || '0');
                    row.push(studentInMonth.summaries[2] || '0');
                    row.push(studentInMonth.summaries[3] || '0');
                }
                data.push(row);
            });

            const ws = (window as any).XLSX.utils.aoa_to_sheet(data);
            ws['!dir'] = 'rtl';
            
            let sheetName = `${st.family} ${st.name}`.replace(/[\[\]\*\/\\\?\:]/g, '').substring(0, 31);
            let baseName = sheetName;
            let counter = 1;
            while (wb.SheetNames.includes(sheetName)) {
                sheetName = `${baseName.substring(0, 28)}_${counter}`;
                counter++;
            }

            (window as any).XLSX.utils.book_append_sheet(wb, ws, sheetName);
        });

        (window as any).XLSX.writeFile(wb, `الأوراق_الفردية_${state.activeClass}.xlsx`);
        showToast('تم تحميل الأوراق الفردية بنجاح', 'success');
    } catch (err) {
        console.error(err);
        showToast('حدث خطأ أثناء تصدير الملف', 'error');
    }
};

let currentSummaryHtml = '';

(window as any).closeSummaryModal = () => {
    const m = document.getElementById('summary-modal');
    if (m) m.classList.remove('open');
};

(window as any).printClassAbsenceSummary = () => {
    const classDatasets = getActiveClassDatasets();
    if (!classDatasets.length) return;

    const firstDs = classDatasets[0];
    const students = firstDs.students;

    if (!students || !students.length) {
        showToast('لا يوجد تلاميذ في هذا القسم', 'error');
        return;
    }

    let h = `<div id="summary-print-content" class="absence-sheet" style="box-shadow:none; margin:0; padding:10mm; width:100%; min-height:100vh; background:white;">
        <div class="sheet-hdr">
            <h2>المملكة المغربية</h2>
            <h3>وزارة التربية الوطنية والتعليم الأولي والرياضة</h3>
            <h3>الأكاديمية الجهوية للتربية والتكوين — ${firstDs.metadata.academy||'—'}</h3>
        </div>
        <div class="info-grid" style="margin-bottom: 20px;">
            <div class="info-cell"><span class="lbl">المؤسسة:</span><span class="val">${firstDs.metadata.institution||'—'}</span></div>
            <div class="info-cell"><span class="lbl">السنة الدراسية:</span><span class="val">${firstDs.metadata.year||'—'}</span></div>
            <div class="info-cell"><span class="lbl">المستوى:</span><span class="val">${firstDs.metadata.level||'—'}</span></div>
            <div class="info-cell"><span class="lbl">القسم:</span><span class="val">${firstDs.metadata.class||'—'}</span></div>
        </div>
        <h3 style="text-align:center; font-weight:bold; font-size:16px; margin-bottom:15px; text-decoration:underline;">لائحة الغياب الإجمالية للتلاميذ</h3>
        <table class="sheet-tbl" style="width:100%;">
            <thead>
                <tr>
                    <th style="width:30px">ت</th>
                    <th style="width:75px">الرقم الوطني</th>
                    <th style="width:35%">الإسم الكامل</th>`;
    
    classDatasets.forEach((ds: any) => {
        h += `<th>${ds.metadata.monthAr || ds.metadata.month}</th>`;
    });
    
    h += `<th style="background:#e5e7eb;">المجموع العام</th>
                </tr>
            </thead>
            <tbody>`;

    students.forEach((st: any) => {
        h += `<tr>
            <td style="font-weight:bold;">${st.rank}</td>
            <td>${st.id}</td>
            <td style="text-align:right; padding-right:8px; font-weight:bold;">${st.family} ${st.name}</td>`;
        
        let grandTotal = 0;
        classDatasets.forEach((ds: any) => {
            const studentInMonth = ds.students.find((s: any) => s.id === st.id);
            const total = studentInMonth ? calcTotalAbsences(studentInMonth) : 0;
            grandTotal += total;
            h += `<td style="${total > 0 ? 'color:#dc2626; font-weight:bold;' : 'color:#059669;'}">${total}</td>`;
        });

        h += `<td style="background:#f9fafb; font-weight:bold; ${grandTotal > 0 ? 'color:#dc2626;' : 'color:#059669;'}">${grandTotal}</td>
        </tr>`;
    });

    h += `</tbody></table>
        <div class="ft-section" style="margin-top: 30px;">
            <div class="ft-row">
                <div class="ft-cell"><span class="fl">توقيع الحارس العام:</span><div class="fl-line"></div></div>
                <div class="ft-cell"><span class="fl">توقيع المدير(ة):</span><div class="fl-line"></div></div>
            </div>
        </div>
    </div>`;

    currentSummaryHtml = h;
    
    const content = document.getElementById('summary-content');
    if (content) content.innerHTML = h;
    
    const m = document.getElementById('summary-modal');
    if (m) m.classList.add('open');
};

(window as any).executePrintSummary = () => {
    const pa = document.getElementById('sheet-print-area');
    if(pa) {
        pa.innerHTML = currentSummaryHtml;
        pa.style.display = 'block';
        (window as any).setPrintOrientation('landscape');
        (window as any).closeSummaryModal();
        setTimeout(() => window.print(), 500);
    }
};

(window as any).downloadSummaryPDF = async () => {
    const content = document.getElementById('summary-print-content');
    if (!content) return;
    
    showToast('جارٍ إنشاء ملف PDF (قد يستغرق بعض الوقت)...', 'info', 10000);
    
    try {
        const { jsPDF } = (window as any).jspdf;
        const pdf = new jsPDF('l', 'mm', 'a4');
        const imgW = 297;
        
        const canvas = await (window as any).html2canvas(content, {
            scale: 2,
            useCORS: true,
            logging: false,
            backgroundColor: '#ffffff'
        });
        
        const imgData = canvas.toDataURL('image/jpeg', 1.0);
        const imgH = (canvas.height * imgW) / canvas.width;
        
        pdf.addImage(imgData, 'JPEG', 0, 0, imgW, imgH);
        pdf.save(`لائحة_الغياب_الإجمالية_${getActiveDataset()?.metadata.class}.pdf`);
        
        showToast('تم تحميل ملف PDF بنجاح', 'success');
    } catch (err) {
        console.error(err);
        showToast('حدث خطأ أثناء إنشاء PDF', 'error');
    }
};

(window as any).exportClassAbsenceListExcel = () => {
    const classDatasets = getActiveClassDatasets();
    if (!classDatasets.length) return;

    const firstDs = classDatasets[0];
    const students = firstDs.students;

    if (!students || !students.length) {
        showToast('لا يوجد تلاميذ في هذا القسم', 'error');
        return;
    }

    const headers = ['ت', 'الرقم الوطني', 'النسب', 'الإسم'];
    classDatasets.forEach((ds: any) => {
        headers.push(ds.metadata.monthAr || ds.metadata.month);
    });
    headers.push('المجموع العام');

    const data = [headers];

    students.forEach((st: any) => {
        const row = [st.rank, st.id, st.family, st.name];
        let grandTotal = 0;

        classDatasets.forEach((ds: any) => {
            const studentInMonth = ds.students.find((s: any) => s.id === st.id);
            const total = studentInMonth ? calcTotalAbsences(studentInMonth) : 0;
            row.push(total);
            grandTotal += total;
        });

        row.push(grandTotal);
        data.push(row);
    });

    try {
        const ws = (window as any).XLSX.utils.aoa_to_sheet(data);
        if(!ws['!cols']) ws['!cols'] = [];
        ws['!dir'] = 'rtl';
        
        const wb = (window as any).XLSX.utils.book_new();
        (window as any).XLSX.utils.book_append_sheet(wb, ws, "لائحة الغياب");
        (window as any).XLSX.writeFile(wb, `لائحة_الغياب_الإجمالية_${state.activeClass}.xlsx`);
        showToast('تم تحميل لائحة الغياب بنجاح', 'success');
    } catch (err) {
        console.error(err);
        showToast('حدث خطأ أثناء تصدير الملف', 'error');
    }
};

(window as any).exportSheetPDF = async () => {
    const sheet = document.getElementById('absence-sheet-combined');
    if (!sheet) return;
    showToast('جارٍ إعداد PDF...', 'info', 5000);
    try {
        const { jsPDF } = (window as any).jspdf;
        const pdf = new jsPDF('l', 'mm', 'a4');
        const imgW = 297;

        const canvas = await (window as any).html2canvas(sheet, {
            scale: 2, useCORS: true, backgroundColor: '#ffffff', logging: false,
            width: sheet.scrollWidth, height: sheet.scrollHeight
        });
        const imgH = (canvas.height * imgW) / canvas.width;
        pdf.addImage(canvas.toDataURL('image/png'), 'PNG', 0, 0, imgW, imgH);

        const ds = getActiveDataset();
        const st = ds?.students.find((s: any) => s.id === currentSheetStudentId);
        const classDatasets = getActiveClassDatasets();
        const monthCount = classDatasets.length;
        pdf.save(`ورقة_غياب_${st ? st.family : ''}_${state.activeClass}${monthCount > 1 ? '_جميع_الأشهر' : '_' + (ds?.metadata.monthAr || '')}.pdf`);
        showToast('تم تصدير PDF بنجاح', 'success');
    } catch (err) {
        console.error(err);
        showToast('خطأ في تصدير PDF', 'error');
    }
};

(window as any).exportAllStudentsSheetsPDF = async () => {
    const classDatasets = getActiveClassDatasets();
    if (!classDatasets.length) return;

    const firstDs = classDatasets[0];
    const students = firstDs.students;

    if (!students || !students.length) {
        showToast('لا يوجد تلاميذ في هذا القسم', 'error');
        return;
    }

    showToast('جارٍ تجهيز الأوراق للتصدير...', 'info');

    let html = '';
    students.forEach((st: any, idx: number) => {
        const sheetId = `absence-sheet-export-all-${idx}`;
        html += buildCombinedAbsenceSheet(classDatasets, st.id, sheetId);
    });

    const pa = document.getElementById('sheet-print-area');
    if (!pa) return;
    
    pa.innerHTML = html;
    pa.style.display = 'block';

    try {
        const sheets = document.querySelectorAll('[id^="absence-sheet-export-all-"]');
        if (!sheets.length) {
            pa.style.display = 'none';
            return;
        }

        showToast('جارٍ إنشاء ملف PDF (قد يستغرق بعض الوقت)...', 'info', 10000);
        const { jsPDF } = (window as any).jspdf;
        const pdf = new jsPDF('l', 'mm', 'a4');
        const imgW = 297;

        for (let i = 0; i < sheets.length; i++) {
            if (i > 0) pdf.addPage();
            const canvas = await (window as any).html2canvas(sheets[i] as HTMLElement, {
                scale: 2, useCORS: true, backgroundColor: '#ffffff', logging: false,
                width: (sheets[i] as HTMLElement).scrollWidth, height: (sheets[i] as HTMLElement).scrollHeight
            });
            const imgH = (canvas.height * imgW) / canvas.width;
            pdf.addImage(canvas.toDataURL('image/png'), 'PNG', 0, 0, imgW, imgH);
        }

        pdf.save(`أوراق_الغياب_جميع_التلاميذ_${state.activeClass}.pdf`);
        showToast('تم تصدير PDF بنجاح', 'success');
    } catch (err) {
        console.error(err);
        showToast('خطأ في تصدير PDF', 'error');
    } finally {
        pa.style.display = 'none';
        pa.innerHTML = '';
    }
};

(window as any).generateClassInvestment = () => {
    const classDatasets = getActiveClassDatasets();
    if (!classDatasets.length) {
        showToast('لا يوجد بيانات لهذا القسم', 'error');
        return;
    }
    
    const m = document.getElementById('ai-modal');
    if (m) m.classList.add('open');
    
    const titleEl = document.getElementById('ai-modal-title');
    const subtitleEl = document.getElementById('ai-modal-subtitle');
    const loadingEl = document.getElementById('ai-loading');
    const contentEl = document.getElementById('ai-content');
    
    if (titleEl) titleEl.innerText = 'استثمار غيابات القسم';
    if (subtitleEl) subtitleEl.innerText = state.activeClass || '';
    if (loadingEl) loadingEl.style.display = 'none';
    
    const firstDs = classDatasets[0];
    const meta = firstDs.metadata;
    
    // Aggregate data
    let totalAbsences = 0;
    let unjustifiedAbsences = 0;
    let justifiedAbsences = 0;
    
    const studentAbsences = new Map<string, { student: any, total: number, unj: number, just: number }>();
    
    classDatasets.forEach((ds: any) => {
        ds.students.forEach((st: any) => {
            const abs = calcTotalAbsences(st);
            const unj = getUnjustified(st);
            const just = getJustified(st);
            
            totalAbsences += abs;
            unjustifiedAbsences += unj;
            justifiedAbsences += just;
            
            if (studentAbsences.has(st.id)) {
                const s = studentAbsences.get(st.id)!;
                s.total += abs;
                s.unj += unj;
                s.just += just;
            } else {
                studentAbsences.set(st.id, { student: st, total: abs, unj: unj, just: just });
            }
        });
    });
    
    const studentsList = Array.from(studentAbsences.values());
    const totalStudents = studentsList.length;
    
    // Sort by total absences
    studentsList.sort((a, b) => b.total - a.total);
    
    const topAbsentees = studentsList.slice(0, 10).filter(s => s.total > 0);
    
    // Distribution
    let zeroAbs = 0, lowAbs = 0, medAbs = 0, highAbs = 0;
    studentsList.forEach(s => {
        if (s.total === 0) zeroAbs++;
        else if (s.total <= 5) lowAbs++;
        else if (s.total <= 10) medAbs++;
        else highAbs++;
    });
    
    let html = `
    <div class="p-6 bg-white text-gray-800" style="direction: rtl; font-family: 'Cairo', sans-serif;">
        <div class="text-center mb-8 border-b pb-4">
            <h2 class="text-2xl font-bold mb-2">استثمار غيابات القسم</h2>
            <div class="flex justify-center gap-6 text-sm font-semibold text-gray-600">
                <span>المؤسسة: ${meta.institution || '—'}</span>
                <span>المستوى: ${meta.level || '—'}</span>
                <span>القسم: ${meta.className || '—'}</span>
                <span>الموسم الدراسي: ${meta.year || '—'}</span>
            </div>
        </div>
        
        <div class="grid grid-cols-2 md:grid-cols-4 gap-4 mb-8">
            <div class="bg-blue-50 p-4 rounded-xl text-center border border-blue-100">
                <p class="text-xs text-blue-600 font-bold mb-1">عدد التلاميذ</p>
                <p class="text-2xl font-black text-blue-800">${totalStudents}</p>
            </div>
            <div class="bg-red-50 p-4 rounded-xl text-center border border-red-100">
                <p class="text-xs text-red-600 font-bold mb-1">مجموع الغيابات</p>
                <p class="text-2xl font-black text-red-800">${totalAbsences}</p>
            </div>
            <div class="bg-orange-50 p-4 rounded-xl text-center border border-orange-100">
                <p class="text-xs text-orange-600 font-bold mb-1">غياب غير مبرر</p>
                <p class="text-2xl font-black text-orange-800">${unjustifiedAbsences}</p>
            </div>
            <div class="bg-emerald-50 p-4 rounded-xl text-center border border-emerald-100">
                <p class="text-xs text-emerald-600 font-bold mb-1">غياب مبرر</p>
                <p class="text-2xl font-black text-emerald-800">${justifiedAbsences}</p>
            </div>
        </div>
        
        <div class="grid grid-cols-1 md:grid-cols-2 gap-8 mb-8">
            <div>
                <h3 class="text-lg font-bold mb-4 border-b pb-2 text-gray-700">توزيع الغيابات</h3>
                <table class="w-full text-sm border-collapse border border-gray-200 text-center">
                    <thead>
                        <tr class="bg-gray-100">
                            <th class="border border-gray-200 p-2">الفئة</th>
                            <th class="border border-gray-200 p-2">عدد التلاميذ</th>
                            <th class="border border-gray-200 p-2">النسبة</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td class="border border-gray-200 p-2 font-semibold text-emerald-600">0 غياب (مواظبون)</td>
                            <td class="border border-gray-200 p-2">${zeroAbs}</td>
                            <td class="border border-gray-200 p-2">${((zeroAbs/totalStudents)*100).toFixed(1)}%</td>
                        </tr>
                        <tr>
                            <td class="border border-gray-200 p-2 font-semibold text-blue-600">1 - 5 غيابات</td>
                            <td class="border border-gray-200 p-2">${lowAbs}</td>
                            <td class="border border-gray-200 p-2">${((lowAbs/totalStudents)*100).toFixed(1)}%</td>
                        </tr>
                        <tr>
                            <td class="border border-gray-200 p-2 font-semibold text-orange-600">6 - 10 غيابات</td>
                            <td class="border border-gray-200 p-2">${medAbs}</td>
                            <td class="border border-gray-200 p-2">${((medAbs/totalStudents)*100).toFixed(1)}%</td>
                        </tr>
                        <tr>
                            <td class="border border-gray-200 p-2 font-semibold text-red-600">أكثر من 10 غيابات</td>
                            <td class="border border-gray-200 p-2">${highAbs}</td>
                            <td class="border border-gray-200 p-2">${((highAbs/totalStudents)*100).toFixed(1)}%</td>
                        </tr>
                    </tbody>
                </table>
            </div>
            
            <div>
                <h3 class="text-lg font-bold mb-4 border-b pb-2 text-gray-700">أكثر التلاميذ غياباً</h3>
                <table class="w-full text-sm border-collapse border border-gray-200 text-center">
                    <thead>
                        <tr class="bg-gray-100">
                            <th class="border border-gray-200 p-2">اسم التلميذ</th>
                            <th class="border border-gray-200 p-2">المجموع</th>
                            <th class="border border-gray-200 p-2">غير مبرر</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${topAbsentees.length > 0 ? topAbsentees.map(s => `
                        <tr>
                            <td class="border border-gray-200 p-2 font-bold text-gray-700">${studentFullName(s.student)}</td>
                            <td class="border border-gray-200 p-2 text-red-600 font-bold">${s.total}</td>
                            <td class="border border-gray-200 p-2 text-orange-600">${s.unj}</td>
                        </tr>
                        `).join('') : `<tr><td colspan="3" class="border border-gray-200 p-4 text-gray-500">لا توجد غيابات مسجلة</td></tr>`}
                    </tbody>
                </table>
            </div>
        </div>
        
        ${classDatasets.length > 1 ? `
        <div class="mb-8">
            <h3 class="text-lg font-bold mb-4 border-b pb-2 text-gray-700">تطور الغيابات حسب الأشهر</h3>
            <table class="w-full text-sm border-collapse border border-gray-200 text-center">
                <thead>
                    <tr class="bg-gray-100">
                        <th class="border border-gray-200 p-2">الشهر</th>
                        <th class="border border-gray-200 p-2">مجموع الغيابات</th>
                        <th class="border border-gray-200 p-2">غير مبرر</th>
                        <th class="border border-gray-200 p-2">متوسط الغياب للتلميذ</th>
                    </tr>
                </thead>
                <tbody>
                    ${classDatasets.map((ds: any) => {
                        let mTotal = 0;
                        let mUnj = 0;
                        ds.students.forEach((st: any) => {
                            mTotal += calcTotalAbsences(st);
                            mUnj += getUnjustified(st);
                        });
                        return `
                        <tr>
                            <td class="border border-gray-200 p-2 font-bold">${ds.metadata.monthAr}</td>
                            <td class="border border-gray-200 p-2 text-red-600 font-bold">${mTotal}</td>
                            <td class="border border-gray-200 p-2 text-orange-600">${mUnj}</td>
                            <td class="border border-gray-200 p-2">${(mTotal / ds.students.length).toFixed(2)}</td>
                        </tr>
                        `;
                    }).join('')}
                </tbody>
            </table>
        </div>
        ` : ''}
        
        <div class="mt-12 pt-8 border-t border-gray-200 flex justify-between text-sm font-bold text-gray-700">
            <div>توقيع الحارس العام:</div>
            <div>توقيع السيد المدير:</div>
        </div>
    </div>
    `;
    
    if (contentEl) contentEl.innerHTML = html;
};

let compareChartInstance: any = null;

(window as any).compareClasses = () => {
    if (!state.datasets || state.datasets.length === 0) {
        showToast('لا توجد بيانات للمقارنة', 'error');
        return;
    }

    const classesData: Record<string, any> = {};
    state.datasets.forEach((ds: any) => {
        const className = ds.metadata.class;
        if (!classesData[className]) {
            classesData[className] = {
                className: className,
                totalAbsences: 0,
                unjustified: 0,
                students: new Set(),
                months: new Set()
            };
        }
        classesData[className].months.add(ds.metadata.monthAr);
        ds.students.forEach((st: any) => {
            classesData[className].students.add(`${st.family} ${st.name}`);
            classesData[className].totalAbsences += calcTotalAbsences(st);
            classesData[className].unjustified += getUnjustified(st);
        });
    });

    const comparisonList = Object.values(classesData).map(c => ({
        className: c.className,
        studentCount: c.students.size,
        totalAbsences: c.totalAbsences,
        unjustified: c.unjustified,
        avgAbsence: c.students.size ? (c.totalAbsences / c.students.size).toFixed(2) : '0',
        months: Array.from(c.months).join('، ')
    }));

    comparisonList.sort((a, b) => b.totalAbsences - a.totalAbsences);

    const tbody = document.getElementById('compare-table-body');
    if (tbody) {
        tbody.innerHTML = comparisonList.map(c => `
            <tr class="border-b border-gray-100 hover:bg-gray-50">
                <td class="py-2 px-3 font-medium text-gray-800">${c.className}</td>
                <td class="py-2 px-3 text-gray-600">${c.studentCount}</td>
                <td class="py-2 px-3 text-red-600 font-bold">${c.totalAbsences}</td>
                <td class="py-2 px-3 text-orange-500">${c.unjustified}</td>
                <td class="py-2 px-3 text-gray-600">${c.avgAbsence}</td>
            </tr>
        `).join('');
    }

    const ctx = document.getElementById('compare-chart') as HTMLCanvasElement;
    if (ctx) {
        if (compareChartInstance) compareChartInstance.destroy();
        compareChartInstance = new (window as any).Chart(ctx, {
            type: 'bar',
            data: {
                labels: comparisonList.map(c => c.className),
                datasets: [
                    {
                        label: 'مجموع الغيابات',
                        data: comparisonList.map(c => c.totalAbsences),
                        backgroundColor: '#ef4444',
                        borderRadius: 4
                    },
                    {
                        label: 'الغيابات غير المبررة',
                        data: comparisonList.map(c => c.unjustified),
                        backgroundColor: '#f97316',
                        borderRadius: 4
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { position: 'top', labels: { font: { family: 'Cairo' } } },
                    tooltip: { titleFont: { family: 'Cairo' }, bodyFont: { family: 'Cairo' } }
                },
                scales: {
                    y: { beginAtZero: true, ticks: { font: { family: 'Cairo' } } },
                    x: { ticks: { font: { family: 'Cairo' } } }
                }
            }
        });
    }

    const aiContent = document.getElementById('compare-ai-content');
    if (aiContent) {
        aiContent.innerHTML = '<div class="text-center text-gray-400 py-4 text-sm">انقر على "توليد استنتاجات" للحصول على قراءة تحليلية للمقارنة.</div>';
    }
    const btnAi = document.getElementById('btn-generate-compare-ai');
    if (btnAi) btnAi.style.display = 'inline-block';

    (window as any).currentComparisonData = comparisonList;

    const m = document.getElementById('compare-modal');
    if (m) m.classList.add('open');
};

(window as any).closeCompareModal = () => {
    const m = document.getElementById('compare-modal');
    if (m) m.classList.remove('open');
};

(window as any).printCompareReport = () => {
    const pa = document.getElementById('sheet-print-area');
    const modalContent = document.getElementById('compare-print-area');
    if (pa && modalContent) {
        const canvas = document.getElementById('compare-chart') as HTMLCanvasElement;
        const chartImg = canvas ? `<img src="${canvas.toDataURL()}" style="max-width: 100%; height: auto; margin-bottom: 20px;">` : '';
        
        const tableHtml = document.querySelector('#compare-print-area .overflow-x-auto')?.outerHTML || '';
        const aiHtml = document.getElementById('compare-ai-content')?.outerHTML || '';
        
        pa.innerHTML = `
            <div style="padding: 20px; font-family: Arial, sans-serif; direction: rtl;">
                <h2 style="text-align: center; margin-bottom: 20px; color: #111;">تقرير مقارنة الأقسام</h2>
                ${tableHtml}
                <div style="margin: 20px 0; text-align: center;">${chartImg}</div>
                <div class="markdown-body" style="direction: rtl; text-align: right;">
                    <h3 style="color: #4f46e5; margin-bottom: 10px;">استنتاجات تحليلية</h3>
                    ${aiHtml}
                </div>
            </div>
        `;
        pa.style.display = 'block';
        (window as any).setPrintOrientation('portrait');
        setTimeout(() => window.print(), 500);
    }
};

(window as any).generateCompareAI = () => {
    const data = (window as any).currentComparisonData;
    if (!data || data.length === 0) return;

    const contentEl = document.getElementById('compare-ai-content');
    const btnAi = document.getElementById('btn-generate-compare-ai');
    
    if (contentEl) contentEl.innerHTML = '';
    if (btnAi) btnAi.style.display = 'none';

    // Sort data
    const sortedByTotal = [...data].sort((a, b) => b.totalAbsences - a.totalAbsences);
    const sortedByAvg = [...data].sort((a, b) => parseFloat(b.avgAbsence) - parseFloat(a.avgAbsence));
    
    const mostAbsentClass = sortedByAvg[0];
    const leastAbsentClass = sortedByAvg[sortedByAvg.length - 1];
    
    const totalAbsencesAll = data.reduce((sum: number, c: any) => sum + c.totalAbsences, 0);
    const totalUnjustifiedAll = data.reduce((sum: number, c: any) => sum + c.unjustified, 0);
    const totalStudentsAll = data.reduce((sum: number, c: any) => sum + c.studentCount, 0);
    
    const overallAvg = totalStudentsAll > 0 ? (totalAbsencesAll / totalStudentsAll).toFixed(2) : '0';
    const unjustifiedRate = totalAbsencesAll > 0 ? ((totalUnjustifiedAll / totalAbsencesAll) * 100).toFixed(1) : '0';

    let html = `
    <div class="p-4 bg-gray-50 rounded-xl border border-gray-200 mt-4 text-gray-800" style="direction: rtl; font-family: 'Cairo', sans-serif;">
        <div class="space-y-4 text-sm leading-relaxed text-justify">
            <p>
                بناءً على المعطيات الإحصائية للأقسام المقارنة، يتبين أن مجموع الغيابات بلغ <strong>${totalAbsencesAll}</strong> وحدة زمنية، 
                منها <strong>${totalUnjustifiedAll}</strong> غياباً غير مبرر (بنسبة <strong>${unjustifiedRate}%</strong>). 
                ويبلغ المعدل العام للغياب <strong>${overallAvg}</strong> وحدة لكل تلميذ.
            </p>
            
            <div class="grid grid-cols-1 md:grid-cols-2 gap-4 mt-3">
                <div class="bg-red-50 p-3 rounded-lg border border-red-100">
                    <h4 class="font-bold text-red-800 mb-1"><i class="fa-solid fa-arrow-trend-up ml-1"></i>الأقسام الأكثر تغيباً</h4>
                    <p>سجل قسم <strong>${mostAbsentClass.className}</strong> أعلى معدل غياب بمتوسط <strong>${mostAbsentClass.avgAbsence}</strong> وحدة لكل تلميذ (مجموع: ${mostAbsentClass.totalAbsences} غياب). يتطلب هذا القسم تدخلاً تربوياً عاجلاً لمعرفة أسباب هذه الظاهرة.</p>
                </div>
                
                <div class="bg-emerald-50 p-3 rounded-lg border border-emerald-100">
                    <h4 class="font-bold text-emerald-800 mb-1"><i class="fa-solid fa-arrow-trend-down ml-1"></i>الأقسام الأكثر انضباطاً</h4>
                    <p>يعتبر قسم <strong>${leastAbsentClass.className}</strong> الأكثر انضباطاً بمعدل غياب بلغ <strong>${leastAbsentClass.avgAbsence}</strong> وحدة لكل تلميذ (مجموع: ${leastAbsentClass.totalAbsences} غياب). يُنصح بتثمين هذا الانضباط وتشجيع تلاميذ القسم.</p>
                </div>
            </div>
            
            <div class="mt-4">
                <h4 class="font-bold text-gray-800 mb-2"><i class="fa-solid fa-lightbulb text-amber-500 ml-1"></i>توصيات عامة:</h4>
                <ul class="list-disc list-inside space-y-1 text-gray-700">
                    <li>تكثيف التواصل مع أولياء أمور تلاميذ الأقسام التي تتصدر لائحة الغياب.</li>
                    <li>التركيز على تقليص نسبة الغياب غير المبرر (${unjustifiedRate}%) من خلال التفعيل الصارم للمذكرات المنظمة.</li>
                    <li>إشراك السادة الأساتذة والموجه التربوي في دراسة أسباب العزوف عن الحضور في الأقسام المتعثرة.</li>
                </ul>
            </div>
        </div>
    </div>
    `;

    if (contentEl) contentEl.innerHTML = html;
};

(window as any).generateAbsenceReport = () => {
    const classDatasets = getActiveClassDatasets();
    if (!classDatasets.length) {
        showToast('لا يوجد بيانات لهذا القسم', 'error');
        return;
    }
    
    const m = document.getElementById('ai-modal');
    if (m) m.classList.add('open');
    
    const titleEl = document.getElementById('ai-modal-title');
    const subtitleEl = document.getElementById('ai-modal-subtitle');
    const loadingEl = document.getElementById('ai-loading');
    const contentEl = document.getElementById('ai-content');
    
    if (titleEl) titleEl.innerText = 'تقرير مفصل حول ظاهرة الغياب';
    if (subtitleEl) subtitleEl.innerText = state.activeClass || '';
    if (loadingEl) loadingEl.style.display = 'none';
    
    const firstDs = classDatasets[0];
    const meta = firstDs.metadata;
    
    // Aggregate data
    let totalAbsences = 0;
    let unjustifiedAbsences = 0;
    
    const studentAbsences = new Map<string, { student: any, total: number, unj: number }>();
    
    classDatasets.forEach((ds: any) => {
        ds.students.forEach((st: any) => {
            const abs = calcTotalAbsences(st);
            const unj = getUnjustified(st);
            
            totalAbsences += abs;
            unjustifiedAbsences += unj;
            
            if (studentAbsences.has(st.id)) {
                const s = studentAbsences.get(st.id)!;
                s.total += abs;
                s.unj += unj;
            } else {
                studentAbsences.set(st.id, { student: st, total: abs, unj: unj });
            }
        });
    });
    
    const studentsList = Array.from(studentAbsences.values());
    const totalStudents = studentsList.length;
    
    // Sort by total absences
    studentsList.sort((a, b) => b.total - a.total);
    
    const topAbsentees = studentsList.slice(0, 5).filter(s => s.total > 0);
    
    const absenceRate = totalStudents > 0 ? (totalAbsences / totalStudents).toFixed(2) : '0';
    const unjustifiedRate = totalAbsences > 0 ? ((unjustifiedAbsences / totalAbsences) * 100).toFixed(1) : '0';
    
    let html = `
    <div class="p-6 bg-white text-gray-800" style="direction: rtl; font-family: 'Cairo', sans-serif;">
        <div class="text-center mb-8 border-b pb-4">
            <h2 class="text-2xl font-bold mb-2">تقرير مفصل حول ظاهرة الغياب وسبل العلاج</h2>
            <div class="flex justify-center gap-6 text-sm font-semibold text-gray-600">
                <span>المؤسسة: ${meta.institution || '—'}</span>
                <span>المستوى: ${meta.level || '—'}</span>
                <span>القسم: ${meta.className || '—'}</span>
                <span>الموسم الدراسي: ${meta.year || '—'}</span>
            </div>
        </div>
        
        <div class="mb-8">
            <h3 class="text-lg font-bold mb-3 text-indigo-700 border-b border-indigo-100 pb-2"><i class="fa-solid fa-magnifying-glass-chart ml-2"></i>1. تشخيص الظاهرة بناءً على الأرقام</h3>
            <p class="mb-4 text-justify leading-relaxed">
                من خلال تحليل بيانات الغياب للقسم <strong>${meta.className || '—'}</strong>، يتبين أن مجموع الغيابات المسجلة بلغ <strong>${totalAbsences}</strong> وحدة زمنية، 
                منها <strong>${unjustifiedAbsences}</strong> غياباً غير مبرر، وهو ما يمثل نسبة <strong>${unjustifiedRate}%</strong> من إجمالي الغيابات. 
                ويبلغ متوسط الغياب لكل تلميذ حوالي <strong>${absenceRate}</strong> وحدة زمنية.
            </p>
            ${topAbsentees.length > 0 ? `
            <p class="mb-2 font-semibold">التلاميذ الأكثر تغيباً والذين يحتاجون تدخلاً عاجلاً:</p>
            <ul class="list-disc list-inside mb-4 text-gray-700 bg-gray-50 p-4 rounded-lg">
                ${topAbsentees.map(s => `<li><strong>${studentFullName(s.student)}</strong>: ${s.total} غياب (منها ${s.unj} غير مبرر).</li>`).join('')}
            </ul>
            ` : '<p class="text-emerald-600 font-semibold mb-4">لا توجد غيابات مسجلة في هذا القسم، مما يدل على انضباط ممتاز.</p>'}
        </div>
        
        <div class="mb-8">
            <h3 class="text-lg font-bold mb-3 text-orange-700 border-b border-orange-100 pb-2"><i class="fa-solid fa-circle-question ml-2"></i>2. الأسباب المحتملة للغياب</h3>
            <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div class="bg-orange-50 p-4 rounded-lg border border-orange-100">
                    <h4 class="font-bold text-orange-800 mb-2">أسباب تربوية</h4>
                    <ul class="list-disc list-inside text-sm text-gray-700 space-y-1">
                        <li>صعوبة مسايرة الإيقاع الدراسي.</li>
                        <li>ضعف التوجيه والمواكبة.</li>
                        <li>الخوف من التقويمات والامتحانات.</li>
                        <li>عدم إنجاز الواجبات المدرسية.</li>
                    </ul>
                </div>
                <div class="bg-blue-50 p-4 rounded-lg border border-blue-100">
                    <h4 class="font-bold text-blue-800 mb-2">أسباب اجتماعية وأسرية</h4>
                    <ul class="list-disc list-inside text-sm text-gray-700 space-y-1">
                        <li>ضعف المراقبة الأسرية.</li>
                        <li>مشاكل عائلية أو اجتماعية.</li>
                        <li>بعد السكن عن المؤسسة.</li>
                        <li>تأثير أصدقاء السوء.</li>
                    </ul>
                </div>
                <div class="bg-purple-50 p-4 rounded-lg border border-purple-100">
                    <h4 class="font-bold text-purple-800 mb-2">أسباب نفسية وصحية</h4>
                    <ul class="list-disc list-inside text-sm text-gray-700 space-y-1">
                        <li>أمراض موسمية أو مزمنة.</li>
                        <li>ضعف الثقة بالنفس.</li>
                        <li>صعوبات في الاندماج مع الزملاء.</li>
                        <li>الشعور بالإحباط أو غياب الدافعية.</li>
                    </ul>
                </div>
            </div>
        </div>
        
        <div class="mb-8">
            <h3 class="text-lg font-bold mb-3 text-emerald-700 border-b border-emerald-100 pb-2"><i class="fa-solid fa-hand-holding-medical ml-2"></i>3. سبل العلاج والتدخلات المقترحة</h3>
            <div class="space-y-3 text-gray-700 text-justify leading-relaxed">
                <p><strong><i class="fa-solid fa-user-tie text-emerald-600 ml-1"></i> دور الإدارة التربوية:</strong> التفعيل الصارم لمقتضيات النظام الداخلي للمؤسسة، استدعاء أولياء أمور التلاميذ الأكثر تغيباً (المذكورين أعلاه) لتوقيع التزامات، وتفعيل مجالس الأقسام لدراسة الحالات المستعصية.</p>
                <p><strong><i class="fa-solid fa-chalkboard-user text-emerald-600 ml-1"></i> دور الأساتذة:</strong> تحفيز التلاميذ، تنويع طرق التدريس لتفادي الملل، رصد الحالات التي تعاني من صعوبات التعلم وإحالتها على الدعم، والتواصل المستمر مع الإدارة بخصوص المتغيبين.</p>
                <p><strong><i class="fa-solid fa-house-chimney-user text-emerald-600 ml-1"></i> دور الأسرة:</strong> ضرورة تتبع مواظبة الأبناء، التواصل المستمر مع إدارة المؤسسة، تبرير الغيابات في آجالها القانونية، وتوفير الجو المناسب للتحصيل بالمنزل.</p>
                <p><strong><i class="fa-solid fa-user-doctor text-emerald-600 ml-1"></i> دور الموجه التربوي:</strong> برمجة جلسات استماع وتوجيه للتلاميذ الذين يعانون من تعثرات دراسية أو نفسية، ومساعدتهم على بناء مشروع شخصي يعزز دافعيتهم للتعلم.</p>
            </div>
        </div>
        
        <div class="mt-12 pt-8 border-t border-gray-200 flex justify-between text-sm font-bold text-gray-700">
            <div>توقيع الحارس العام:</div>
            <div>توقيع السيد المدير:</div>
        </div>
    </div>
    `;
    
    if (contentEl) contentEl.innerHTML = html;
};

(window as any).closeAiModal = () => {
    const m = document.getElementById('ai-modal');
    if (m) m.classList.remove('open');
};

(window as any).printAiReport = () => {
    const content = document.getElementById('ai-content')?.innerHTML || '';
    const pa = document.getElementById('sheet-print-area');
    if (pa) {
        pa.innerHTML = `<div style="padding: 20px; font-family: 'Cairo', sans-serif; direction: rtl;">${content}</div>`;
        pa.style.display = 'block';
        (window as any).setPrintOrientation('portrait');
        setTimeout(() => window.print(), 500);
    }
};

(window as any).promptForApiKey = () => {
    const key = prompt('الرجاء إدخال مفتاح API الخاص بك (Kimi أو OpenAI):');
    if (key && key.trim() !== '') {
        localStorage.setItem('custom_ai_api_key', key.trim());
        showToast('تم حفظ المفتاح بنجاح. يرجى إعادة المحاولة.', 'success');
        // Close modal or retry
        const modal = document.getElementById('ai-modal');
        if (modal) modal.classList.remove('active');
        const compareModal = document.getElementById('compare-modal');
        if (compareModal) compareModal.classList.remove('active');
    }
};

(window as any).selectApiKeyAndRetryReport = async () => {
    if ((window as any).aistudio) {
        try {
            await (window as any).aistudio.openSelectKey();
            (window as any).generateAbsenceReport();
        } catch (e) {
            console.error('Failed to select API key', e);
        }
    } else {
        showToast('ميزة اختيار المفتاح غير متوفرة في هذه البيئة', 'error');
    }
};

(window as any).printReport = () => {
    const ds = getActiveDataset(); if (!ds) return; buildPrintArea(ds, null); 
    (window as any).setPrintOrientation('portrait');
    window.print();
};

(window as any).exportPDF = async () => {
    const ds = getActiveDataset(); if (!ds) return;
    showToast('جارٍ إعداد PDF...', 'info', 5000);
    try {
        buildPrintArea(ds, null);
        const pa = document.getElementById('print-area');
        if(!pa) return;
        const canvas = await (window as any).html2canvas(pa, { scale: 2, useCORS: true, backgroundColor: '#ffffff', logging: false, width: pa.scrollWidth, height: pa.scrollHeight });
        const { jsPDF } = (window as any).jspdf; const imgW = 297, pageH = 210;
        const imgH = (canvas.height * imgW) / canvas.width;
        const pdf = new jsPDF('l', 'mm', 'a4');
        let hL = imgH, pos = 0;
        pdf.addImage(canvas.toDataURL('image/png'), 'PNG', 0, pos, imgW, imgH); hL -= pageH;
        while (hL > 0) { pos -= pageH; pdf.addPage(); pdf.addImage(canvas.toDataURL('image/png'), 'PNG', 0, pos, imgW, imgH); hL -= pageH; }
        pdf.save(`غيابات_${ds.metadata.class}_${ds.metadata.monthAr}.pdf`);
        showToast('تم تصدير PDF', 'success');
    } catch (err) { console.error(err); showToast('خطأ في PDF', 'error'); }
};

let confirmCallback: (() => void) | null = null;

(window as any).showConfirmModal = (message: string, callback: () => void) => {
    const m = document.getElementById('confirm-modal');
    const msgEl = document.getElementById('confirm-modal-message');
    if (m && msgEl) {
        msgEl.textContent = message;
        confirmCallback = callback;
        m.classList.add('open');
    }
};

(window as any).openGuideModal = () => {
    const m = document.getElementById('guide-modal');
    if (m) m.classList.add('open');
};

(window as any).closeGuideModal = () => {
    const m = document.getElementById('guide-modal');
    if (m) m.classList.remove('open');
};

(window as any).closeConfirmModal = () => {
    const m = document.getElementById('confirm-modal');
    if (m) m.classList.remove('open');
    confirmCallback = null;
};

document.getElementById('confirm-modal-yes')?.addEventListener('click', () => {
    if (confirmCallback) confirmCallback();
    (window as any).closeConfirmModal();
});

(window as any).removeActiveClass = async () => {
    if (!state.activeClass) return;
    
    const classDatasets = getActiveClassDatasets();
    if (!classDatasets.length) return;

    (window as any).showConfirmModal(`هل أنت متأكد من حذف جميع بيانات القسم "${state.activeClass}"؟`, async () => {
        showToast('جارٍ حذف القسم...', 'info');
        
        // Remove from Firebase if logged in
        if (currentUser) {
            for (const ds of classDatasets) {
                const datasetId = getDatasetId(ds);
                await deleteDatasetFromFirebase(datasetId);
            }
        }
        
        // Remove from local state
        state.datasets = state.datasets.filter((ds: any) => ds.metadata.class !== state.activeClass);
        
        // Reset active class
        const remainingClasses = Array.from(getUniqueClasses().keys());
        if (remainingClasses.length > 0) {
            state.activeClass = remainingClasses[0];
            const cm = getClassMonths(state.activeClass);
            state.activeMonth = cm.length ? cm[0].metadata.month : null;
        } else {
            state.activeClass = null;
            state.activeMonth = null;
        }
        
        state.searchQuery = '';
        const si = document.getElementById('search-input') as HTMLInputElement; 
        if (si) si.value = '';
        
        [barChart, lineChart, pieChart, hBarChart].forEach(c => { if (c) { c.destroy(); c = null; } });
        
        saveDataLocally();
        renderAll();
        showToast('تم حذف القسم بنجاح', 'success');
    });
};

(window as any).clearAllData = async () => {
    if(!state.datasets.length)return;
    
    (window as any).showConfirmModal('هل أنت متأكد من حذف جميع البيانات؟ سيتم حذفها من قاعدة البيانات أيضاً.', async () => {
        if (currentUser) {
            showToast('جارٍ الحذف...', 'info');
            for (const ds of state.datasets) {
                const datasetId = getDatasetId(ds);
                await deleteDatasetFromFirebase(datasetId);
            }
        }
        
        state.datasets=[];state.activeClass=null;state.activeMonth=null;state.searchQuery='';state.latenessRecords=[];
        const si = document.getElementById('search-input') as HTMLInputElement; if(si) si.value='';
        [barChart,lineChart,pieChart,hBarChart].forEach(c=>{if(c){c.destroy();c=null}});
        saveDataLocally();
        renderAll();
        showToast('تم حذف جميع البيانات','info');
    });
};

(window as any).setBarChartMode = (m: string) => {
    state.barChartMode=m;
    const bt = document.getElementById('bar-mode-total'); if(bt) bt.classList.toggle('active',m==='total');
    const bu = document.getElementById('bar-mode-unjustified'); if(bu) bu.classList.toggle('active',m==='unjustified');
    updateBarChart();
};

(window as any).updateLineChart = updateLineChart;

(window as any).setViewMode = (m: string) => {
    state.viewMode=m;
    const vc = document.getElementById('view-compact'); if(vc) vc.classList.toggle('active',m==='compact');
    const vd = document.getElementById('view-detailed'); if(vd) vd.classList.toggle('active',m==='detailed');
    renderTable();
};

(window as any).setSortMode = (m: string) => {
    state.sortMode=m;
    const sr = document.getElementById('sort-rank'); if(sr) sr.classList.toggle('active',m==='rank');
    const sa = document.getElementById('sort-absent'); if(sa) sa.classList.toggle('active',m==='absent');
    renderTable();
};

(window as any).handleSearch = (q: string) => {
    state.searchQuery=q.trim();renderTable();
};

(window as any).openAbsenceSheet = (sid: string) => {
    const classDatasets = getActiveClassDatasets();
    if (!classDatasets.length) return;

    const firstSt = classDatasets[0].students.find((s: any) => s.id === sid);
    if (!firstSt) return;

    currentSheetStudentId = sid;
    const monthCount = classDatasets.length;

    const sms = document.getElementById('sheet-modal-subtitle');
    if(sms) sms.textContent = `${studentFullName(firstSt)} — ${state.activeClass}`;

    const countBadge = document.getElementById('sheet-months-count');
    if (countBadge) {
        if (monthCount > 1) {
            countBadge.style.display = '';
            countBadge.textContent = `${monthCount} أشهر`;
        } else {
            countBadge.style.display = 'none';
        }
    }

    let html = buildCombinedAbsenceSheet(classDatasets, sid, 'absence-sheet-combined');

    const sp = document.getElementById('sheet-preview'); if(sp) sp.innerHTML = html;
    const sm = document.getElementById('sheet-modal'); if(sm) sm.classList.add('open');
};

(window as any).openDetail = (sid: string) => {
    const ds=getActiveDataset();if(!ds)return;const st=ds.students.find((s: any)=>s.id===sid);if(!st)return;
    const sd=getSchoolDays(ds.students),ta=calcTotalAbsences(st),ad=calcAbsentDays(st);
    
    // Lateness records for this student
    const studentLateness = state.latenessRecords.filter((r: any) => r.studentId === sid).sort((a: any, b: any) => new Date(b.date).getTime() - new Date(a.date).getTime());
    let latenessHtml = '';
    if (studentLateness.length > 0) {
        latenessHtml = `<div class="mb-6"><h4 class="text-sm font-bold text-gray-700 mb-3"><i class="fa-solid fa-clock text-rose-600 ml-2"></i>التأخرات (${studentLateness.length})</h4><div class="card p-4">`;
        studentLateness.forEach((r: any) => {
            latenessHtml += `<div class="flex items-center justify-between py-2 border-b border-gray-50 last:border-0"><div class="flex flex-col"><span class="text-sm font-bold text-gray-800">${r.date}</span><span class="text-xs text-gray-500">${r.subject}</span></div><span class="px-3 py-1 rounded-lg text-xs font-bold bg-rose-100 text-rose-700">${r.minutes} دقيقة</span></div>`;
        });
        latenessHtml += `</div></div>`;
    }

    let calH='';['أحد','إثنين','ثلاثاء','أربعاء','خميس','جمعة','سبت'].forEach(d=>{calH+=`<div class="text-center text-xs text-gray-400 font-semibold py-1">${d}</div>`});
    for(let d=1;d<=31;d++){const info=parseDayValue(st.days[d]);let bg='bg-gray-50 text-gray-300';if(info.type==='present')bg='bg-emerald-50 text-emerald-600';else if(info.type==='absent')bg=info.val>=2?'bg-red-500 text-white':'bg-red-100 text-red-600';else if(info.type==='special')bg='bg-amber-100 text-amber-700';calH+=`<div class="${bg} ${!sd.has(d)?'opacity-40':''}" style="aspect-ratio:1;border-radius:10px;display:flex;flex-direction:column;align-items:center;justify-content:center;font-size:.72rem;font-weight:600"><span style="font-size:.7rem">${d}</span>${info.type==='absent'?`<span style="font-size:.6rem;font-weight:700">${info.val}</span>`:''}</div>`}
    let absH='';for(let d=1;d<=31;d++){const info=parseDayValue(st.days[d]);if(info.type==='absent'||info.type==='special'){absH+=`<div class="flex items-center justify-between py-2 border-b border-gray-50"><span class="text-sm font-semibold">اليوم ${d}</span><span class="px-3 py-1 rounded-lg text-xs font-bold ${info.type==='special'?'bg-amber-100 text-amber-700':'bg-red-100 text-red-600'}">${info.type==='special'?'X':info.val+' غياب'}</span></div>`}}
    const pc = document.getElementById('panel-content');
    if(pc) {
        pc.innerHTML=`
        <div class="flex items-center justify-between mb-6"><h3 class="text-lg font-extrabold text-gray-800">تفاصيل التلميذ</h3><button onclick="closeDetailPanel()" class="btn-ghost rounded-xl w-9 h-9 flex items-center justify-center"><i class="fa-solid fa-xmark text-gray-400"></i></button></div>
        <div class="bg-gradient-to-l from-emerald-700 to-emerald-800 rounded-2xl p-5 text-white mb-6"><div class="flex items-center gap-4 mb-4"><div class="w-14 h-14 rounded-xl bg-white/15 flex items-center justify-center text-2xl font-extrabold">${st.family.charAt(0)}</div><div><p class="text-lg font-extrabold">${st.family} ${st.name}</p><p class="text-emerald-200 text-sm font-light">${st.id}</p></div></div><div class="grid grid-cols-3 gap-3"><div class="bg-white/10 rounded-xl p-3 text-center"><p class="text-xl font-extrabold">${ad}</p><p class="text-xs text-emerald-200">أيام غياب</p></div><div class="bg-white/10 rounded-xl p-3 text-center"><p class="text-xl font-extrabold">${ta}</p><p class="text-xs text-emerald-200">مجموع</p></div><div class="bg-white/10 rounded-xl p-3 text-center"><p class="text-xl font-extrabold">${sd.size-ad}</p><p class="text-xs text-emerald-200">حضور</p></div></div></div>
        ${getActiveClassDatasets().length>1?`<div class="mb-6"><h4 class="text-sm font-bold text-gray-700 mb-3"><i class="fa-solid fa-chart-line text-amber-600 ml-2"></i>تطور الغياب</h4><div class="card p-4" style="max-height:220px"><canvas id="detailLineChart"></canvas></div></div>`:''}
        <div class="mb-6"><h4 class="text-sm font-bold text-gray-700 mb-3"><i class="fa-solid fa-calendar-days text-emerald-600 ml-2"></i>تقويم — ${ds.metadata.monthAr}</h4><div class="card p-4"><div style="display:grid;grid-template-columns:repeat(7,1fr);gap:6px">${calH}</div></div></div>
        ${absH?`<div class="mb-6"><h4 class="text-sm font-bold text-gray-700 mb-3"><i class="fa-solid fa-list-check text-red-500 ml-2"></i>أيام الغياب</h4><div class="card p-4">${absH}</div></div>`:`<div class="text-center py-8"><i class="fa-solid fa-circle-check text-4xl text-emerald-300 mb-3"></i><p class="text-sm text-gray-400">لا غيابات</p></div>`}
        ${latenessHtml}
        ${ds.summaryCols.length>0?`<div class="mb-6"><div class="card p-4"><div class="grid grid-cols-2 gap-3">${['يوم غير مبرر','يوم مبرر','ساعة غير مبررة','ساعة مبررة'].map((l,i)=>`<div class="bg-gray-50 rounded-xl p-3 text-center"><p class="text-lg font-extrabold ${parseInt(st.summaries[i])>0?'text-red-600':'text-gray-400'}">${st.summaries[i]||'0'}</p><p class="text-xs text-gray-500">${l}</p></div>`).join('')}</div></div></div>`:''}
        <div class="flex gap-3 mt-2"><button onclick="closeDetailPanel();openAbsenceSheet('${st.id}')" class="btn btn-primary flex-1 justify-center"><i class="fa-solid fa-file-lines"></i> ورقة الغياب</button><button onclick="printStudentReport('${st.id}')" class="btn btn-outline flex-1 justify-center"><i class="fa-solid fa-print"></i> طباعة التقرير</button></div>`;
    }
    const dp = document.getElementById('detail-panel'); if(dp) dp.classList.add('open');
    const po = document.getElementById('panel-overlay'); if(po) po.classList.add('open');
    if(getActiveClassDatasets().length>1){setTimeout(()=>{const cds=getActiveClassDatasets();const canvas = document.getElementById('detailLineChart') as HTMLCanvasElement; if(canvas) { new (window as any).Chart(canvas,{type:'line',data:{labels:cds.map((d: any)=>d.metadata.monthAr),datasets:[{label:'غياب',data:cds.map((d: any)=>{const s=d.students.find((x: any)=>x.id===sid);return s?calcTotalAbsences(s):0}),borderColor:'#D97706',backgroundColor:'rgba(217,119,6,.1)',fill:true,tension:.4,pointRadius:5,pointHoverRadius:8,borderWidth:3}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{rtl:true,titleFont:{family:'Cairo'},bodyFont:{family:'Cairo'}}},scales:{x:{ticks:{font:{family:'Cairo',size:10}},grid:{color:'#f0f1f4'}},y:{beginAtZero:true,ticks:{font:{family:'Cairo',size:10},stepSize:1},grid:{color:'#f0f1f4'}}}}}) }},100)}
};

(window as any).printStudentReport = (sid: string) => {
    const ds = getActiveDataset(); if (!ds) return; const st = ds.students.find((s: any) => s.id === sid); if (!st) return; (window as any).closeDetailPanel(); buildPrintArea(ds, st); 
    (window as any).setPrintOrientation('portrait');
    setTimeout(() => window.print(), 200);
};

// Setup event listeners
document.addEventListener('DOMContentLoaded', () => {
    const uploadZone=document.getElementById('upload-zone');
    const fileInput=document.getElementById('file-input') as HTMLInputElement;
    const guardianFileInput=document.getElementById('guardian-file-input') as HTMLInputElement;
    if(uploadZone && fileInput) {
        ['dragenter','dragover'].forEach(e=>uploadZone.addEventListener(e,ev=>{ev.preventDefault();uploadZone.classList.add('drag-over')}));
        ['dragleave','drop'].forEach(e=>uploadZone.addEventListener(e,ev=>{ev.preventDefault();uploadZone.classList.remove('drag-over')}));
        uploadZone.addEventListener('drop',e=>handleFiles(e.dataTransfer!.files));
        fileInput.addEventListener('change',e=>{handleFiles((e.target as HTMLInputElement).files!); (e.target as HTMLInputElement).value=''});
    }
    if (guardianFileInput) {
        guardianFileInput.addEventListener('change', e => {
            handleGuardianFiles((e.target as HTMLInputElement).files!);
            (e.target as HTMLInputElement).value = '';
        });
    }

    function handleFiles(files: FileList){Array.from(files).forEach(file=>{if(!file.name.match(/\.(xlsx|xls)$/i)){showToast(`"${file.name}" ليس Excel`,'error');return}if(file.size>MAX_FILE_SIZE){showToast(`"${file.name}" كبير جداً`,'error');return}const r=new FileReader();r.onload=e=>processExcel(e.target!.result,file.name);r.onerror=()=>showToast(`خطأ في القراءة`,'error');r.readAsArrayBuffer(file)})}

    function handleGuardianFiles(files: FileList) {
        Array.from(files).forEach(file => {
            if (!file.name.match(/\.(xlsx|xls)$/i)) {
                showToast(`"${file.name}" ليس Excel`, 'error');
                return;
            }
            const r = new FileReader();
            r.onload = e => processGuardianExcel(e.target!.result);
            r.onerror = () => showToast(`خطأ في قراءة ملف أولياء الأمور`, 'error');
            r.readAsArrayBuffer(file);
        });
    }

    document.addEventListener('keydown', e => { if (e.key === 'Escape') { (window as any).closeDetailPanel(); (window as any).closeSheetModal(); } });
    window.addEventListener('afterprint', () => { 
        const pa = document.getElementById('print-area'); if(pa) pa.style.display = 'none'; 
        const spa = document.getElementById('sheet-print-area'); if(spa) spa.style.display = 'none'; 
        const styleEl = document.getElementById('print-orientation-style');
        if (styleEl) styleEl.remove();
    });
    
    // Auth listeners
    const btnLogin = document.getElementById('btn-login');
    const btnLogout = document.getElementById('btn-logout');
    
    // Initial load from local storage
    loadDataLocally();
    renderAll();
    
    if (btnLogin) {
        btnLogin.addEventListener('click', async () => {
            try {
                await signInWithPopup(auth, googleProvider);
            } catch (error) {
                console.error("Login error", error);
                showToast('فشل تسجيل الدخول. تم تفعيل الحفظ المحلي التلقائي.', 'info', 5000);
                const authWarning = document.getElementById('auth-warning');
                if (authWarning) authWarning.style.display = 'none';
            }
        });
    }
    
    if (btnLogout) {
        btnLogout.addEventListener('click', async () => {
            try {
                await signOut(auth);
                state.datasets = [];
                state.activeClass = null;
                state.activeMonth = null;
                renderAll();
            } catch (error) {
                console.error("Logout error", error);
            }
        });
    }
    
    onAuthStateChanged(auth, async (user) => {
        currentUser = user;
        const btnLogin = document.getElementById('btn-login');
        const userInfo = document.getElementById('user-info');
        const userName = document.getElementById('user-name');
        const userAvatar = document.getElementById('user-avatar') as HTMLImageElement;
        const authWarning = document.getElementById('auth-warning');
        
        if (user) {
            if (btnLogin) btnLogin.style.display = 'none';
            if (userInfo) userInfo.style.display = 'flex';
            if (userName) userName.textContent = user.displayName || 'مستخدم';
            if (userAvatar) userAvatar.src = user.photoURL || '';
            if (authWarning) authWarning.style.display = 'none';
            
            await loadGuardiansFromFirebase();
            await loadDatasetsFromFirebase();
        } else {
            if (btnLogin) btnLogin.style.display = 'inline-flex';
            if (userInfo) userInfo.style.display = 'none';
            if (authWarning) {
                if (state.datasets.length > 0) {
                    authWarning.style.display = 'none'; // Hide warning if local data exists
                } else {
                    authWarning.style.display = 'flex';
                }
            }
        }
    });
    
    renderAll();
});

