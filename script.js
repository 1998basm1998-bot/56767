// ============ 1. قاعدة البيانات والتهيئة ============
let db = JSON.parse(localStorage.getItem('schoolDB_ERP')) || {
    schoolName: '', schoolDate: '', theme: 'light',
    classes: [], regions: [], students: [],
    staff: [], expenses: [], dynamicSheets: {} 
};

let currentStudentId = null, editingStudentId = null, tempSiblings = [];
const defaultSubjects = ['العربي','الرياضيات','الإنكليزي','الإسلامية','الكيمياء','الفيزياء','الاحياء','الفنية','الرياضة','الجرائم'];

// ============ 2. محرك التنبيهات 3D (SweetAlert2) ============
function customAlert(msg, icon = 'info') {
    Swal.fire({ text: msg, icon: icon, confirmButtonText: 'موافق', customClass: { popup: 'swal2-glass' }});
}
function customConfirm(msg, cb) {
    Swal.fire({ text: msg, icon: 'warning', showCancelButton: true, confirmButtonText: 'نعم', cancelButtonText: 'إلغاء', customClass: { popup: 'swal2-glass' }
    }).then((res) => { cb(res.isConfirmed); });
}
function customPrompt(msg, cb) {
    Swal.fire({ title: msg, input: 'text', showCancelButton: true, confirmButtonText: 'تأكيد', cancelButtonText: 'إلغاء', customClass: { popup: 'swal2-glass' }
    }).then((res) => { cb(res.isConfirmed ? res.value : null); });
}

window.onload = () => {
    applyTheme(db.theme);
    if (!db.schoolName) { document.getElementById('setup-modal').classList.add('active'); } 
    else { document.getElementById('setup-modal').classList.remove('active'); document.getElementById('app').classList.remove('hidden'); initApp(); }
};

function saveDB() { localStorage.setItem('schoolDB_ERP', JSON.stringify(db)); }

function saveSetup() {
    let name = document.getElementById('setup-school-name').value, date = document.getElementById('setup-school-date').value;
    if (!name) return customAlert('يرجى كتابة اسم المدرسة', 'warning');
    db.schoolName = name; db.schoolDate = date; saveDB(); location.reload();
}

function initApp() {
    if(!db.staff) db.staff = []; if(!db.expenses) db.expenses = []; if(!db.dynamicSheets) db.dynamicSheets = {};
    document.getElementById('display-school-name').innerHTML = `<i class="fas fa-university"></i> ${db.schoolName}`;
    document.getElementById('edit-school-name').value = db.schoolName; document.getElementById('edit-school-date').value = db.schoolDate;
    
    renderClasses(); renderRegions(); renderStudents(); populateClassSelects(); updateRepSectionDropdown();
    renderDynamicTabs();
}

function updateSchool() {
    db.schoolName = document.getElementById('edit-school-name').value; db.schoolDate = document.getElementById('edit-school-date').value;
    saveDB(); initApp(); customAlert('تم التحديث بنجاح', 'success');
}

function switchTab(tabId) {
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active-tab'));
    document.querySelectorAll('.nav-btn').forEach(el => el.classList.remove('active'));
    let tab = document.getElementById(`tab-${tabId}`), nav = document.getElementById(`nav-${tabId}`);
    if(tab) tab.classList.add('active-tab'); if(nav) nav.classList.add('active');
    
    if(tabId === 'reports') updateRepSectionDropdown();
    if(tabId === 'salaries') renderStaff();
    if(tabId === 'daily') { document.getElementById('daily-date-filter').value = new Date().toISOString().split('T')[0]; renderDaily(); }
    if(tabId === 'dual') renderDual();
}

function toggleTheme() { db.theme = db.theme === 'light' ? 'dark' : 'light'; applyTheme(db.theme); saveDB(); }
function applyTheme(theme) { document.body.classList.toggle('dark-mode', theme === 'dark'); }
function showModal(id) { document.getElementById(id).classList.add('active'); }
function hideModal(id) { document.getElementById(id).classList.remove('active'); }

// ============ 3. محرك ERP الشامل (تحويل Excel إلى تبويبات ويب) ============
function importExcel() {
    let fileInput = document.getElementById('excel-file');
    if (!fileInput.files[0]) return customAlert('يرجى اختيار ملف', 'warning');
    
    Swal.fire({ title: 'جاري التحليل...', allowOutsideClick: false, didOpen: () => Swal.showLoading(), customClass:{popup:'swal2-glass'} });

    let reader = new FileReader();
    reader.onload = function(e) {
        try {
            let data = new Uint8Array(e.target.result);
            let workbook = XLSX.read(data, {type: 'array'});
            if (!db.dynamicSheets) db.dynamicSheets = {};
            
            workbook.SheetNames.forEach(sheetName => {
                let json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {defval: ""});
                if (json.length > 0) {
                    let cols = [], firstRow = json[0];
                    for (let key in firstRow) {
                        let type = 'text', isNum = json.some(r => r[key] !== "" && !isNaN(r[key]));
                        if (isNum) type = 'number';
                        if (key.toLowerCase().includes('date') || key.includes('تاريخ')) type = 'date';
                        cols.push({ name: key, type: type });
                    }
                    
                    let records = json.map(row => {
                        let rec = { id: Date.now() + Math.random() };
                        cols.forEach(col => { rec[col.name] = col.type === 'number' && row[col.name] !== "" ? parseFloat(row[col.name]) : row[col.name]; });
                        return rec;
                    });
                    
                    db.dynamicSheets[sheetName] = { cols: cols, data: records };
                }
            });
            saveDB(); renderDynamicTabs(); fileInput.value = '';
            Swal.fire({icon: 'success', title: 'تمت العملية', text: 'تم إنشاء التبويبات والمحاسبة الديناميكية', customClass:{popup:'swal2-glass'}});
        } catch(err) { customAlert('تأكد من صيغة الملف وعدم وجود أخطاء فيه.', 'error'); }
    };
    reader.readAsArrayBuffer(fileInput.files[0]);
}

function renderDynamicTabs() {
    document.querySelectorAll('.dynamic-tab').forEach(el => el.remove());
    document.querySelectorAll('.dynamic-nav').forEach(el => el.remove());
    
    let bottomNav = document.getElementById('bottom-nav-container'), app = document.getElementById('app');
    if (!db.dynamicSheets) return;
    
    Object.keys(db.dynamicSheets).forEach((sheetName, index) => {
        let tabId = 'dyn-' + index;
        
        let btn = document.createElement('button');
        btn.className = 'nav-btn dynamic-nav'; btn.id = 'nav-' + tabId;
        btn.onclick = () => switchTab(tabId);
        btn.innerHTML = `<i class="fas fa-file-excel"></i><span>${sheetName.substring(0, 8)}</span>`;
        bottomNav.appendChild(btn);
        
        let tab = document.createElement('div');
        tab.className = 'tab-content dynamic-tab pb-x'; tab.id = 'tab-' + tabId;
        tab.innerHTML = `
            <div class="glass p-3 mb-3">
                <div class="flex-between mb-2">
                    <h3 style="color:#2ecc71; margin:0;">${sheetName}</h3>
                    <button class="btn-3d success btn-small m-0" onclick="openDynModal('${sheetName}')"><i class="fas fa-plus"></i> إدراج</button>
                </div>
                <div style="position:relative;"><i class="fas fa-search" style="position:absolute; right:15px; top:15px; color:#888;"></i><input type="text" class="glass-input m-0" style="padding-right:40px;" placeholder="بحث..." onkeyup="renderDynList('${sheetName}', '${tabId}', this.value)"></div>
                <div id="dyn-sums-${tabId}" class="grid-2 mt-2"></div>
            </div>
            <div id="dyn-list-${tabId}" class="list-container"></div>
        `;
        app.insertBefore(tab, bottomNav); renderDynList(sheetName, tabId, "");
    });
}

function renderDynList(sheetName, tabId, filter) {
    let sheet = db.dynamicSheets[sheetName]; if (!sheet) return;
    let listEl = document.getElementById(`dyn-list-${tabId}`), sumsEl = document.getElementById(`dyn-sums-${tabId}`);
    
    let html = '', sums = {};
    sheet.cols.forEach(c => { if(c.type === 'number') sums[c.name] = 0; });
    
    sheet.data.forEach(row => {
        let match = filter === "" || Object.values(row).some(v => String(v).includes(filter));
        if (match) {
            sheet.cols.forEach(c => { if(c.type === 'number' && !isNaN(row[c.name])) sums[c.name] += parseFloat(row[c.name]); });
            let cardHtml = sheet.cols.map(c => `<div><small style="opacity:0.7;">${c.name}:</small> <b>${row[c.name]||'-'}</b></div>`).join('');
            html += `<div class="list-item"><div class="grid-2" style="font-size:14px;">${cardHtml}</div><div class="flex-row mt-2" style="justify-content:flex-end;"><button class="btn-3d warning btn-small m-0" onclick="openDynModal('${sheetName}', ${row.id})"><i class="fas fa-pen"></i></button><button class="btn-3d danger btn-small m-0" onclick="deleteDynRecord('${sheetName}', ${row.id})"><i class="fas fa-trash"></i></button></div></div>`;
        }
    });
    listEl.innerHTML = html || '<div class="text-center">لا توجد بيانات</div>';
    sumsEl.innerHTML = Object.keys(sums).map(k => `<div class="glass p-2 text-center">إجمالي ${k}<br><b class="text-success" style="font-size:18px;">${sums[k].toLocaleString()}</b></div>`).join('');
}

let currentDynSheet = null, currentDynRecordId = null;
function openDynModal(sheetName, recordId = null) {
    currentDynSheet = sheetName; currentDynRecordId = recordId;
    let sheet = db.dynamicSheets[sheetName], record = recordId ? sheet.data.find(r => r.id === recordId) : null;
    document.getElementById('dyn-modal-title').innerHTML = recordId ? '<i class="fas fa-edit"></i> تعديل' : '<i class="fas fa-plus"></i> إضافة';
    
    let html = sheet.cols.map((c, i) => {
        let val = record ? (record[c.name] || '') : '';
        let inputType = c.type === 'number' ? 'number' : (c.type === 'date' ? 'date' : 'text');
        if (c.type === 'text') {
            let uniqueVals = [...new Set(sheet.data.map(r => r[c.name]))].filter(v => v);
            if (uniqueVals.length > 0 && uniqueVals.length <= 15) {
                return `<div><label style="font-size:14px;">${c.name}</label><select id="dyn_inp_${i}" class="glass-input m-0 mt-1"><option value="">اختر...</option>${uniqueVals.map(v => `<option value="${v}" ${v==val?'selected':''}>${v}</option>`).join('')}</select></div>`;
            }
        }
        return `<div><label style="font-size:14px;">${c.name}</label><input type="${inputType}" id="dyn_inp_${i}" class="glass-input m-0 mt-1" value="${val}"></div>`;
    }).join('');
    
    document.getElementById('dyn-modal-inputs').innerHTML = html; showModal('dynamic-modal');
}

function saveDynRecord() {
    let sheet = db.dynamicSheets[currentDynSheet], rec = { id: currentDynRecordId || Date.now() };
    sheet.cols.forEach((c, i) => { let val = document.getElementById(`dyn_inp_${i}`).value; rec[c.name] = c.type === 'number' ? (parseFloat(val) || 0) : val; });
    if (currentDynRecordId) { let idx = sheet.data.findIndex(r => r.id === currentDynRecordId); sheet.data[idx] = rec; } else { sheet.data.push(rec); }
    saveDB(); hideModal('dynamic-modal');
    let tabIndex = Object.keys(db.dynamicSheets).indexOf(currentDynSheet);
    renderDynList(currentDynSheet, `dyn-${tabIndex}`, "");
    Swal.fire({toast:true, position:'top-end', icon:'success', title:'تم الحفظ', showConfirmButton:false, timer:1500});
}
function deleteDynRecord(sheetName, id) {
    customConfirm("تأكيد الحذف نهائياً؟", (res) => {
        if (res) { db.dynamicSheets[sheetName].data = db.dynamicSheets[sheetName].data.filter(r => r.id !== id); saveDB();
            let tabIndex = Object.keys(db.dynamicSheets).indexOf(sheetName); renderDynList(sheetName, `dyn-${tabIndex}`, ""); }
    });
}

// ============ 4. الوحدات الأصلية المبرمجة سابقاً (الطلاب والمناطق) ============
function addClass() { let n=document.getElementById('new-class-name').value; if(!n)return; db.classes.push({id:Date.now(), name:n, sections:[]}); saveDB(); document.getElementById('new-class-name').value=''; initApp(); }
function deleteClass(id) { customConfirm('تأكيد حذف الصف بالكامل؟', r=>{ if(r){db.classes=db.classes.filter(c=>c.id!==id); saveDB(); initApp();}}); }
function addSection(id) { customPrompt("اكتب اسم الشعبة:", n=>{ if(n){db.classes.find(c=>c.id===id).sections.push({id:Date.now(),name:n}); saveDB(); initApp();}});}
function deleteSection(cId, sId) { customConfirm('حذف الشعبة؟', r=>{ if(r){let c=db.classes.find(x=>x.id===cId); c.sections=c.sections.filter(x=>x.id!==sId); saveDB(); initApp();}});}
function renderClasses() {
    document.getElementById('classes-list').innerHTML = db.classes.map(c => `<div class="list-item"><div class="flex-between w-100 mb-1"><strong>${c.name}</strong> <div><button class="btn-3d success btn-small m-0" onclick="addSection(${c.id})"><i class="fas fa-plus"></i> شعبة</button> <button class="btn-3d danger btn-small m-0" onclick="deleteClass(${c.id})"><i class="fas fa-trash"></i></button></div></div><div>${c.sections.map(s => `<span class="glass p-1" style="display:inline-block; margin:2px; cursor:pointer;" ondblclick="deleteSection(${c.id},${s.id})">${s.name} ✖</span>`).join('')}</div></div>`).join('');
}
function addRegion() {
    let n=document.getElementById('reg-name').value, d=document.getElementById('reg-driver').value, p=document.getElementById('reg-phone').value, c=document.getElementById('reg-code').value;
    if(!n||!c) return customAlert("الاسم والرمز مطلوبان", 'warning');
    db.regions.push({id:Date.now(), name:n, driver:d, phone:p, code:c}); saveDB(); renderRegions(); ['reg-name','reg-driver','reg-phone','reg-code'].forEach(id=>document.getElementById(id).value='');
}
function deleteRegion(id) { customConfirm('تأكيد الحذف؟', r=>{ if(r){db.regions=db.regions.filter(x=>x.id!==id); saveDB(); renderRegions();} });}
function renderRegions() { document.getElementById('regions-list').innerHTML = db.regions.map(r => `<div class="list-item flex-between"><div><b>${r.name}</b> (الرمز: ${r.code})<br><small><i class="fas fa-bus"></i> السائق: ${r.driver} | ${r.phone}</small></div><button class="btn-3d danger btn-small m-0" onclick="deleteRegion(${r.id})"><i class="fas fa-trash"></i></button></div>`).join(''); }

function populateClassSelects() { let opts = '<option value="">اختر الصف</option>'+db.classes.map(c=>`<option value="${c.id}">${c.name}</option>`).join(''); ['std-class','sib-class','rep-class'].forEach(id=>document.getElementById(id).innerHTML=opts); }
function updateSectionDropdown() { updateDropdown('std-class', 'std-section'); } function updateSibSectionDropdown() { updateDropdown('sib-class', 'sib-section'); } function updateRepSectionDropdown() { updateDropdown('rep-class', 'rep-section'); generateReport(); }
function updateDropdown(cId, sId, sVal=null) { let v=document.getElementById(cId).value, sel=document.getElementById(sId); sel.innerHTML='<option value="">اختر الشعبة</option>'; let c=db.classes.find(x=>x.id==v); if(c){ sel.innerHTML+=c.sections.map(s=>`<option value="${s.id}">${s.name}</option>`).join(''); if(sVal) sel.value=sVal;} }
function checkRegionCode() { let reg=db.regions.find(r=>r.code===document.getElementById('std-reg-code').value); document.getElementById('std-reg-name').value=reg?reg.name:''; document.getElementById('std-driver-name').value=reg?reg.driver:''; document.getElementById('std-driver-phone').value=reg?reg.phone:''; }

function openAddStudentModal() { editingStudentId=null; tempSiblings=[]; renderTempSiblings(); ['std-class','std-section','std-reg','std-name','std-phone','std-fee','std-reg-code','std-reg-name','std-driver-name','std-driver-phone'].forEach(id=>document.getElementById(id).value=''); showModal('add-student-modal'); }
function editStudent(id) { editingStudentId=id; let s=db.students.find(x=>x.id===id); document.getElementById('std-class').value=s.classId; updateDropdown('std-class','std-section',s.sectionId); document.getElementById('std-reg').value=s.regId||''; document.getElementById('std-name').value=s.name; document.getElementById('std-phone').value=s.phone||''; document.getElementById('std-fee').value=s.tuition; document.getElementById('std-reg-code').value=s.regionCode||''; checkRegionCode(); tempSiblings=s.siblings?JSON.parse(JSON.stringify(s.siblings)):[]; renderTempSiblings(); showModal('add-student-modal'); }
function saveSiblingTemp() { let n=document.getElementById('sib-name').value; if(!n)return customAlert('اسم الأخ مطلوب', 'warning'); tempSiblings.push({regId:document.getElementById('sib-reg').value, name:n, classId:document.getElementById('sib-class').value, sectionId:document.getElementById('sib-section').value}); renderTempSiblings(); hideModal('add-sibling-modal'); ['sib-reg','sib-name','sib-class','sib-section'].forEach(id=>document.getElementById(id).value=''); }
function renderTempSiblings() { document.getElementById('siblings-temp-list').innerHTML=tempSiblings.map((s,i)=>`<div class="list-item flex-between p-2"><span><i class="fas fa-child"></i> ${s.name}</span> <button class="btn-3d danger btn-small m-0" onclick="tempSiblings.splice(${i},1); renderTempSiblings()"><i class="fas fa-trash"></i></button></div>`).join(''); }

function saveStudent() {
    let n=document.getElementById('std-name').value; if(!n) return customAlert('الاسم الثلاثي مطلوب','error');
    let d = { classId:document.getElementById('std-class').value, sectionId:document.getElementById('std-section').value, regId:document.getElementById('std-reg').value, name:n, phone:document.getElementById('std-phone').value, tuition:parseFloat(document.getElementById('std-fee').value)||0, regionCode:document.getElementById('std-reg-code').value, siblings:[...tempSiblings] };
    if(editingStudentId){ let idx=db.students.findIndex(x=>x.id===editingStudentId); d.id=db.students[idx].id; d.payments=db.students[idx].payments; d.grades=db.students[idx].grades; db.students[idx]=d; Swal.fire({toast:true,position:'top-end',icon:'success',title:'تم التعديل',showConfirmButton:false,timer:1500}); } 
    else { d.id=Date.now(); d.payments=[]; d.grades={}; db.students.push(d); Swal.fire({toast:true,position:'top-end',icon:'success',title:'تم الإضافة',showConfirmButton:false,timer:1500}); }
    saveDB(); hideModal('add-student-modal'); renderStudents();
}

function searchStudent() { renderStudents(document.getElementById('search-student').value); }
function renderStudents(filter="") {
    let html='';
    db.students.forEach(s=>{
        let matchM=s.name.includes(filter) || (s.regId && s.regId.includes(filter)); let matchS=(s.siblings||[]).find(sib=>sib.name.includes(filter));
        if(matchM || (filter!=="" && matchS)) {
            let cls=db.classes.find(c=>c.id==s.classId), paid=s.payments.reduce((sum,p)=>sum+p.amount,0);
            html+=`<div class="list-item flex-between" style="cursor:pointer;" onclick="openProfile(${s.id})"><div><strong><i class="fas fa-user-graduate text-primary"></i> ${s.name}</strong><br><small>الصف: ${cls?cls.name:'-'} | الباقي: <span style="color:#e74c3c">${s.tuition-paid}</span></small></div><div class="flex-row"><button class="btn-3d warning btn-small m-0" onclick="event.stopPropagation(); editStudent(${s.id})"><i class="fas fa-pen"></i></button><button class="btn-3d danger btn-small m-0" onclick="event.stopPropagation(); deleteStudent(${s.id})"><i class="fas fa-trash"></i></button></div></div>`;
        }
        if(filter==="" || matchM || matchS){
            (s.siblings||[]).forEach(sib=>{
                if(filter==="" || sib.name.includes(filter) || matchM){
                    let sc=db.classes.find(c=>c.id==sib.classId);
                    html+=`<div class="list-item flex-between" style="cursor:pointer; background:rgba(0,0,0,0.05); border-right:4px solid #a777e3;" onclick="openProfile(${s.id})"><div><strong><i class="fas fa-child"></i> ${sib.name}</strong> <small style="color:#a777e3;">(أخو ${s.name})</small><br><small>الصف: ${sc?sc.name:''}</small></div><button class="btn-3d btn-small m-0" style="background:#a777e3;" onclick="event.stopPropagation(); openProfile(${s.id})"><i class="fas fa-folder-open"></i> الحساب</button></div>`;
                }
            });
        }
    }); document.getElementById('students-list').innerHTML=html;
}
function deleteStudent(id) { customConfirm("تأكيد الحذف نهائياً؟", r=>{ if(r){db.students=db.students.filter(s=>s.id!==id); saveDB(); renderStudents();} }); }

// ============ 5. التقارير المالية (Native ERP) ============
// الرواتب
function openAddStaffModal() { document.getElementById('staff-name').value=''; document.getElementById('staff-role').value=''; document.getElementById('staff-salary').value=''; showModal('add-staff-modal'); }
function saveStaff() {
    let name=document.getElementById('staff-name').value, role=document.getElementById('staff-role').value, salary=parseFloat(document.getElementById('staff-salary').value)||0;
    if(!name || salary<=0) return customAlert('الاسم والراتب مطلوبان', 'error');
    db.staff.push({ id:Date.now(), name, role, salary, payments:[] }); saveDB(); hideModal('add-staff-modal'); renderStaff();
}
function renderStaff() {
    if(!db.staff) return;
    document.getElementById('staff-list').innerHTML = db.staff.map(st => {
        let paid = st.payments.reduce((s, p) => s + p.amount, 0);
        return `<div class="list-item flex-between"><div><strong><i class="fas fa-user-tie text-primary"></i> ${st.name}</strong> <small>(${st.role})</small><br><small>الراتب: ${st.salary} | المدفوع: <span style="color:#2ecc71">${paid}</span></small></div><div class="flex-row"><button class="btn-3d success btn-small m-0" onclick="payStaff(${st.id})"><i class="fas fa-hand-holding-usd"></i> صرف</button><button class="btn-3d danger btn-small m-0" onclick="deleteStaff(${st.id})"><i class="fas fa-trash"></i></button></div></div>`;
    }).join('');
}
function payStaff(id) {
    customPrompt("المبلغ المراد صرفه للموظف:", (amt) => {
        let val = parseFloat(amt);
        if(val > 0) { let st = db.staff.find(x => x.id === id); st.payments.push({ amount: val, date: new Date().toISOString().split('T')[0] }); saveDB(); renderStaff(); Swal.fire({toast:true, position:'top-end', icon:'success', title:'تم الصرف', showConfirmButton:false, timer:1500}); }
    });
}
function deleteStaff(id) { customConfirm("حذف الموظف وسجلاته؟", (res) => { if(res){ db.staff = db.staff.filter(x => x.id !== id); saveDB(); renderStaff(); } }); }

// خلاصة اليوم
function renderDaily() {
    let dVal = document.getElementById('daily-date-filter').value;
    let parts = dVal.split('-'); let dStrGB = parts.length === 3 ? `${parts[2]}/${parts[1]}/${parts[0]}` : ''; // لدعم التنسيق القديم
    
    let totalIn = 0, totalOut = 0;
    db.students.forEach(s => s.payments.forEach(p => { if(p.date === dVal || p.date === dStrGB) totalIn += p.amount; }));
    if(db.staff) db.staff.forEach(st => st.payments.forEach(p => { if(p.date === dVal || p.date === dStrGB) totalOut += p.amount; }));
    if(db.expenses) db.expenses.forEach(e => { if(e.date === dVal || e.date === dStrGB) totalOut += e.amount; });
    
    document.getElementById('daily-in').innerText = totalIn.toLocaleString(); document.getElementById('daily-out').innerText = totalOut.toLocaleString(); document.getElementById('daily-net').innerText = (totalIn - totalOut).toLocaleString();
    if(!db.expenses) db.expenses = [];
    document.getElementById('daily-expenses').innerHTML = db.expenses.filter(e => e.date === dVal || e.date === dStrGB).map(e => `<div class="list-item flex-between p-2"><span>${e.desc}</span> <b>${e.amount}</b></div>`).join('');
}
function openAddExpense() {
    customPrompt("اكتب الوصف ثم المبلغ (مثال: صيانة 50000):", (val) => {
        if(val) {
            let parts = val.split(' '), amt = parseFloat(parts[parts.length-1]), desc = parts.slice(0, parts.length-1).join(' ');
            if(!isNaN(amt) && desc) {
                let dVal = document.getElementById('daily-date-filter').value;
                db.expenses.push({ id: Date.now(), desc, amount: amt, date: dVal }); saveDB(); renderDaily();
            } else { customAlert("صيغة الإدخال خاطئة. اكتب الوصف وبعده مسافة ثم المبلغ.", 'error'); }
        }
    });
}

// الخلاصة الثنائية
function renderDual() {
    let exp = 0, col = 0, sal = 0, exs = 0;
    db.students.forEach(s => { exp += s.tuition; col += s.payments.reduce((a,b)=>a+b.amount, 0); });
    if(db.staff) db.staff.forEach(st => sal += st.payments.reduce((a,b)=>a+b.amount, 0));
    if(db.expenses) db.expenses.forEach(e => exs += e.amount);
    
    document.getElementById('dual-expected').innerText = exp.toLocaleString(); document.getElementById('dual-collected').innerText = col.toLocaleString();
    document.getElementById('dual-debt').innerText = (exp - col).toLocaleString(); document.getElementById('dual-salaries').innerText = sal.toLocaleString();
    document.getElementById('dual-expenses').innerText = exs.toLocaleString();
    
    let net = col - sal - exs;
    let netEl = document.getElementById('dual-net'); netEl.innerText = net.toLocaleString(); netEl.style.color = net >= 0 ? '#2ecc71' : '#e74c3c';
}

// ============ 6. عمليات الطالب العادية والدرجات والتقارير ============
function generateReport() {
    let cId=document.getElementById('rep-class').value, sId=document.getElementById('rep-section').value; if(!cId||!sId){document.getElementById('report-list').innerHTML='';return;}
    let f=db.students.filter(s=>s.classId==cId && s.sectionId==sId).map(s=>({...s, isSib:false}));
    db.students.forEach(m=>{ (m.siblings||[]).forEach(sib=>{ if(sib.classId==cId && sib.sectionId==sId) f.push({name:sib.name, isSib:true, mName:m.name}); }); });
    f.sort((a,b)=>a.name.localeCompare(b.name,'ar'));
    document.getElementById('report-list').innerHTML=f.map((s,i)=>`<div class="list-item"><b>${i+1}.</b> ${s.name} ${s.isSib?`<small style="color:red;">(أخ لـ ${s.mName})</small>`:''}</div>`).join('');
}
function exportReportExcel() {
    let cId=document.getElementById('rep-class').value, sId=document.getElementById('rep-section').value; let c=db.classes.find(x=>x.id==cId), s=c?c.sections.find(x=>x.id==sId):null;
    if(!c||!s) return customAlert('اختر الصف والشعبة', 'warning');
    let f=db.students.filter(x=>x.classId==cId && x.sectionId==sId).map(x=>({...x, isSib:false, ph:x.phone}));
    db.students.forEach(m=>{ (m.siblings||[]).forEach(sib=>{ if(sib.classId==cId && sib.sectionId==sId) f.push({name:sib.name, isSib:true, ph:m.phone}); }); });
    f.sort((a,b)=>a.name.localeCompare(b.name,'ar'));
    let data=[[`تقرير الصف: ${c.name} - الشعبة: ${s.name}`],["التسلسل","اسم الطالب","ملاحظة","الموبايل"]];
    f.forEach((x,i)=>{ data.push([i+1, x.name, x.isSib?'أخ/أخت':'', x.ph]); });
    let wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), "التقرير"); XLSX.writeFile(wb, `تقرير_${c.name}_${s.name}.xlsx`);
}

function openProfile(id) {
    currentStudentId=id; let s=db.students.find(x=>x.id===id); let c=db.classes.find(x=>x.id==s.classId), sec=c?c.sections.find(x=>x.id==s.sectionId):null;
    document.getElementById('prof-name').innerText=s.name; document.getElementById('prof-details').innerText=`الصف: ${c?c.name:'-'} | الشعبة: ${sec?sec.name:'-'} | موبايل: ${s.phone||'-'}`; document.getElementById('prof-reg').innerText=s.regId||'-';
    updateFinance(s); showModal('student-profile-modal');
}
function updateFinance(s) {
    let paid=s.payments.reduce((sum,p)=>sum+p.amount,0); document.getElementById('prof-total').innerText=s.tuition; document.getElementById('prof-paid').innerText=paid; document.getElementById('prof-rem').innerText=s.tuition-paid;
    document.getElementById('payments-history').innerHTML=s.payments.map((p,i)=>`<div class="list-item flex-between p-2"><span>${p.amount} د.ع | ${p.date}</span> <button class="btn-3d danger btn-small m-0" onclick="delPayment(${i})"><i class="fas fa-trash"></i></button></div>`).join('');
}
function submitPayment() { let amt=parseFloat(document.getElementById('pay-amount').value); if(amt>0){ let s=db.students.find(x=>x.id===currentStudentId); s.payments.push({amount:amt, date:new Date().toLocaleDateString('en-GB')}); saveDB(); document.getElementById('pay-amount').value=''; updateFinance(s); Swal.fire({toast:true, position:'top-end', icon:'success', title:'تم التسديد', showConfirmButton:false, timer:1500}); } }
function delPayment(i) { customConfirm('حذف التسديد؟', r=>{ if(r){let s=db.students.find(x=>x.id===currentStudentId); s.payments.splice(i,1); saveDB(); updateFinance(s);} }); }

function openGrades() {
    let s=db.students.find(x=>x.id===currentStudentId); if(!s.grades)s.grades={}; if(!s.grades.subNames)s.grades.subNames=[...defaultSubjects];
    let html=`<table class="grades-tbl"><thead><tr class="bg-light"><th rowspan="2" class="bg-blue">المادة</th><th colspan="3">الفصل الأول</th><th rowspan="2" class="bg-yellow">معدل ف1</th><th rowspan="2">نصف السنة</th><th colspan="3">الفصل الثاني</th><th rowspan="2" class="bg-yellow">معدل ف2</th><th rowspan="2" class="bg-yellow">السعي السنوي</th><th rowspan="2">الامتحان النهائي</th><th rowspan="2" class="bg-yellow">الدرجة النهائية</th></tr><tr class="bg-light"><th>ش1</th><th>ش2</th><th>ش3</th><th>ش1</th><th>ش2</th><th>ش3</th></tr></thead><tbody>`;
    for(let i=0; i<10; i++){ let g=s.grades[i]||{}; html+=`<tr><td><input type="text" id="g_sub_${i}" value="${s.grades.subNames[i]||''}" class="sub-name" oninput="calcCols()"></td><td><input type="number" id="g_${i}_m11" value="${g.m11||''}" oninput="calcRow(${i});calcCols();"></td><td><input type="number" id="g_${i}_m12" value="${g.m12||''}" oninput="calcRow(${i});calcCols();"></td><td><input type="number" id="g_${i}_m13" value="${g.m13||''}" oninput="calcRow(${i});calcCols();"></td><td class="bg-yellow"><input type="number" id="g_${i}_avg1" value="${g.avg1||''}" readonly></td><td><input type="number" id="g_${i}_mid" value="${g.mid||''}" oninput="calcRow(${i});calcCols();"></td><td><input type="number" id="g_${i}_m21" value="${g.m21||''}" oninput="calcRow(${i});calcCols();"></td><td><input type="number" id="g_${i}_m22" value="${g.m22||''}" oninput="calcRow(${i});calcCols();"></td><td><input type="number" id="g_${i}_m23" value="${g.m23||''}" oninput="calcRow(${i});calcCols();"></td><td class="bg-yellow"><input type="number" id="g_${i}_avg2" value="${g.avg2||''}" readonly></td><td class="bg-yellow"><input type="number" id="g_${i}_year" value="${g.year||''}" readonly></td><td><input type="number" id="g_${i}_final" value="${g.final||''}" oninput="calcRow(${i});calcCols();"></td><td class="bg-yellow"><input type="number" id="g_${i}_tot" value="${g.tot||''}" readonly></td></tr>`; }
    html+=`<tr class="bg-light"><td class="bg-yellow">المجموع</td><td class="bg-yellow"><input type="number" id="g_tot_m11" readonly></td><td class="bg-yellow"><input type="number" id="g_tot_m12" readonly></td><td class="bg-yellow"><input type="number" id="g_tot_m13" readonly></td><td class="bg-yellow"><input type="number" id="g_tot_avg1" readonly></td><td class="bg-yellow"><input type="number" id="g_tot_mid" readonly></td><td class="bg-yellow"><input type="number" id="g_tot_m21" readonly></td><td class="bg-yellow"><input type="number" id="g_tot_m22" readonly></td><td class="bg-yellow"><input type="number" id="g_tot_m23" readonly></td><td class="bg-yellow"><input type="number" id="g_tot_avg2" readonly></td><td class="bg-yellow"><input type="number" id="g_tot_year" readonly></td><td class="bg-yellow"><input type="number" id="g_tot_final" readonly></td><td class="bg-yellow"><input type="number" id="g_tot_tot" readonly></td></tr><tr class="bg-light"><td class="bg-yellow">المعدل</td><td class="bg-yellow"><input type="number" id="g_avg_m11" readonly></td><td class="bg-yellow"><input type="number" id="g_avg_m12" readonly></td><td class="bg-yellow"><input type="number" id="g_avg_m13" readonly></td><td class="bg-yellow"><input type="number" id="g_avg_avg1" readonly></td><td class="bg-yellow"><input type="number" id="g_avg_mid" readonly></td><td class="bg-yellow"><input type="number" id="g_avg_m21" readonly></td><td class="bg-yellow"><input type="number" id="g_avg_m22" readonly></td><td class="bg-yellow"><input type="number" id="g_avg_m23" readonly></td><td class="bg-yellow"><input type="number" id="g_avg_avg2" readonly></td><td class="bg-yellow"><input type="number" id="g_avg_year" readonly></td><td class="bg-yellow"><input type="number" id="g_avg_final" readonly></td><td class="bg-yellow"><input type="number" id="g_avg_tot" readonly></td></tr><tr><td>النتيجة</td><td colspan="12"><input type="text" id="g_footer" value="${s.grades['footer']||''}" style="width:100%; text-align:right;" placeholder="اكتب النتيجة هنا..."></td></tr></tbody></table>`;
    document.getElementById('grades-table-container').innerHTML=html; for(let i=0;i<10;i++)calcRow(i); calcCols(); showModal('grades-modal');
}
function calcRow(i) {
    let get = id => {let v=parseFloat(document.getElementById(id).value); return isNaN(v)?null:v;}; let set = (id,v) => document.getElementById(id).value = v!==null?Math.round(v):''; let avg = arr => {let f=arr.filter(x=>x!==null); return f.length?f.reduce((a,b)=>a+b)/f.length:null;};
    let a1=avg([get(`g_${i}_m11`),get(`g_${i}_m12`),get(`g_${i}_m13`)]); set(`g_${i}_avg1`,a1);
    let a2=avg([get(`g_${i}_m21`),get(`g_${i}_m22`),get(`g_${i}_m23`)]); set(`g_${i}_avg2`,a2);
    let yr=avg([a1,get(`g_${i}_mid`),a2]); set(`g_${i}_year`,yr);
    let fn=get(`g_${i}_final`), tot=(yr!==null&&fn!==null)?(yr+fn)/2:null; set(`g_${i}_tot`,tot);
}
function calcCols() {
    let cols=['m11','m12','m13','avg1','mid','m21','m22','m23','avg2','year','final','tot'], vC=Array.from({length:10}).filter((_,i)=>document.getElementById(`g_sub_${i}`).value.trim()!=='').length||1;
    cols.forEach(c => { let sum=0, cnt=0; for(let i=0;i<10;i++){let v=parseFloat(document.getElementById(`g_${i}_${c}`).value); if(!isNaN(v)){sum+=v;cnt++;}} document.getElementById(`g_tot_${c}`).value=cnt>0?Math.round(sum):''; document.getElementById(`g_avg_${c}`).value=cnt>0?(sum/vC).toFixed(1).replace(/\.0$/,''):''; });
}
function saveGrades() {
    let s=db.students.find(x=>x.id===currentStudentId); s.grades.subNames=[];
    for(let i=0;i<10;i++){ s.grades.subNames.push(document.getElementById(`g_sub_${i}`).value); s.grades[i]={m11:document.getElementById(`g_${i}_m11`).value, m12:document.getElementById(`g_${i}_m12`).value, m13:document.getElementById(`g_${i}_m13`).value, avg1:document.getElementById(`g_${i}_avg1`).value, mid:document.getElementById(`g_${i}_mid`).value, m21:document.getElementById(`g_${i}_m21`).value, m22:document.getElementById(`g_${i}_m22`).value, m23:document.getElementById(`g_${i}_m23`).value, avg2:document.getElementById(`g_${i}_avg2`).value, year:document.getElementById(`g_${i}_year`).value, final:document.getElementById(`g_${i}_final`).value, tot:document.getElementById(`g_${i}_tot`).value}; } s.grades['footer']=document.getElementById('g_footer').value; saveDB(); Swal.fire({toast:true,position:'top-end',icon:'success',title:'تم حفظ الدرجات',showConfirmButton:false,timer:1500});
}

function printGrades() { let s=db.students.find(x=>x.id===currentStudentId), c=db.classes.find(x=>x.id==s.classId); document.querySelectorAll('#grades-table-container input').forEach(e=>e.setAttribute('value',e.value)); document.getElementById('print-area').innerHTML=`<div style="font-family:Arial; padding:20px; direction:rtl;"><div style="display:flex; justify-content:space-between; background:#2c3e50; color:#fff; padding:15px; font-weight:bold; font-size:22px; border: 2px solid #000;"><span>${db.schoolName}</span> <span>الصف: ${c?c.name:''}</span></div><div style="text-align:center; background:#f1c40f; padding:15px; font-size:24px; font-weight:bold; border: 2px solid #000; border-top:none; margin-bottom:10px;">الطالب: <span style="color:#e74c3c;">${s.name}</span></div>${document.getElementById('grades-table-container').innerHTML}</div>`; window.print(); }
function printReceipt() {
    let s=db.students.find(x=>x.id===currentStudentId), c=db.classes.find(x=>x.id==s.classId), sc=c?c.sections.find(x=>x.id==s.sectionId):null, p=s.payments.reduce((a,b)=>a+b.amount,0);
    document.getElementById('print-area').innerHTML=`<div style="font-family:Arial; padding:20px; border:2px solid #000; width:95%; margin:auto; direction:rtl;"><div style="display:flex; justify-content:space-between; border-bottom:3px double #000; padding-bottom:15px; margin-bottom:15px;"><div><b>التاريخ:</b><br>${new Date().toLocaleDateString('en-GB')}</div><h1 style="margin:0; color:#2c3e50;">${db.schoolName}</h1><div><b>القيد:</b><br>${s.regId||'----'}</div></div><table style="width:100%; border-collapse:collapse; text-align:center; font-weight:bold; border:2px solid #000; margin-bottom:15px;" border="1"><tr style="background:#ecf0f1; -webkit-print-color-adjust:exact;"><td style="padding:10px;">علاقة</td><td>الاسم</td><td>القيد</td><td>الصف</td><td>الشعبة</td></tr><tr><td style="padding:10px;">الرئيسي</td><td>${s.name}</td><td>${s.regId||'-'}</td><td>${c?c.name:''}</td><td>${sc?sc.name:''}</td></tr>${s.siblings?s.siblings.map(sib=>{let xC=db.classes.find(y=>y.id==sib.classId), xS=xC?xC.sections.find(y=>y.id==sib.sectionId):null; return `<tr><td style="padding:10px; color:#e74c3c;">أخ/أخت</td><td>${sib.name}</td><td>${sib.regId||'-'}</td><td>${xC?xC.name:''}</td><td>${xS?xS.name:''}</td></tr>`;}).join(''):''}</table><table style="width:100%; border-collapse:collapse; text-align:center; font-weight:bold; border:2px solid #000; margin-bottom:15px;" border="1"><tr style="background:#ecf0f1; -webkit-print-color-adjust:exact;"><td style="padding:10px;">المبلغ الكلي</td><td>الواصل (المسدد)</td><td colspan="2">المتبقي (الذمة)</td></tr><tr><td style="padding:10px; font-size:18px;">${s.tuition}</td><td style="font-size:18px;">${p}</td><td colspan="2" style="font-size:18px; color:red;">${s.tuition-p}</td></tr></table><table style="width:100%; border-collapse:collapse; text-align:center; font-weight:bold; border:1px solid #000;" border="1"><tr style="background:#ecf0f1; -webkit-print-color-adjust:exact;"><td colspan="2">الدفعة 1</td><td colspan="2">الدفعة 2</td><td colspan="2">الدفعة 3</td><td colspan="2">الدفعة 4</td></tr><tr><td>المبلغ</td><td>التاريخ</td><td>المبلغ</td><td>التاريخ</td><td>المبلغ</td><td>التاريخ</td><td>المبلغ</td><td>التاريخ</td></tr><tr><td>${s.payments[0]?.amount||0}</td><td>${s.payments[0]?.date||'-'}</td><td>${s.payments[1]?.amount||0}</td><td>${s.payments[1]?.date||'-'}</td><td>${s.payments[2]?.amount||0}</td><td>${s.payments[2]?.date||'-'}</td><td>${s.payments[3]?.amount||0}</td><td>${s.payments[3]?.date||'-'}</td></tr><tr style="background:#ecf0f1; -webkit-print-color-adjust:exact;"><td colspan="2">الدفعة 5</td><td colspan="2">الدفعة 6</td><td colspan="2">الدفعة 7</td><td colspan="2">الدفعة 8</td></tr><tr><td>${s.payments[4]?.amount||0}</td><td>${s.payments[4]?.date||'-'}</td><td>${s.payments[5]?.amount||0}</td><td>${s.payments[5]?.date||'-'}</td><td>${s.payments[6]?.amount||0}</td><td>${s.payments[6]?.date||'-'}</td><td>${s.payments[7]?.amount||0}</td><td>${s.payments[7]?.date||'-'}</td></tr></table><div style="font-size:13px; text-align:center; margin-top:15px; border-top:1px dashed #000; padding-top:10px;">ولي الأمر ملزم بدفع المبلغ كاملاً دون قطع أي مبلغ في حال النقل أو الترك</div></div>`; window.print();
}
function exportStudentExcel() { let s=db.students.find(x=>x.id===currentStudentId), c=db.classes.find(x=>x.id==s.classId), sc=c?c.sections.find(x=>x.id==s.sectionId):null, p=s.payments.reduce((a,b)=>a+b.amount,0); let data=[["العلاقة","الاسم","القيد","الموبايل","الصف","الشعبة","الكلي","الواصل","المتبقي"],["الرئيسي",s.name,s.regId||'',s.phone||'',c?c.name:'',sc?sc.name:'',s.tuition,p,s.tuition-p]]; if(s.siblings){s.siblings.forEach(sib=>{let xC=db.classes.find(y=>y.id==sib.classId), xS=xC?xC.sections.find(y=>y.id==sib.sectionId):null; data.push(["أخ/أخت",sib.name,sib.regId||'',"-",xC?xC.name:'',xS?xS.name:'',"","-","-"]);});} data.push([],["الدفعات","التاريخ","المبلغ"]); s.payments.forEach((p,i)=>data.push([`دفعة ${i+1}`,p.date,p.amount])); let wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(data),"الحساب"); XLSX.writeFile(wb,`الطالب_${s.name}.xlsx`); }
