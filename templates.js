const TEMPLATE_STORES=[['3888','53 Hoàng Bá Bích'],['4044','389 Đường N6'],['4112','38 Đặng Văn Trơn'],['4139','157-159 Phan Đình Phùng'],['4163','3.9 Nguyễn Văn Tỏ'],['4227','869 Hoàng Tam Kỳ'],['4324','A83 đường 5'],['4354','81 Nguyễn Hoàng'],['4465','G1 Khu 94 Long Đức 1'],['4468','152 Đinh Quang Ân'],['4673','27 Phùng Hưng'],['5140','175-177 đường N16'],['2DEL','Nguyễn Hoàng'],['6055','G1/31 Tổ 19'],['6138','Tổ 13 KP Vườn Dừa'],['6390','167 Ngô Quyền'],['6488','1111 Bùi Văn Hòa'],['2BH7','Lô 17-18 KDC Bình Dương']];
const OPS_CHT={
'3014064':'3888','6006212':'4044','6094846':'4112','978428':'4139','6141047':'4163','6217819':'4227','3565221':'4324','3577317':'4354','3577322':'4465','3589208':'4468','6062282':'4673','3587506':'5140','3592078':'2DEL','3606571':'6055','6063420':'6138','957636':'6390','3003399':'6488','3613897':'2BH7'
};
function cleanMNV(v){return String(v??'').replace(/\D/g,'').replace(/^0+/,'')||'0';}
function plain(v){return String(v??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/Đ/g,'D').trim().toUpperCase();}
function fallbackCSV(rows,name){const esc=v=>'"'+String(v??'').replace(/"/g,'""')+'"';const csv='\ufeff'+rows.map(r=>r.map(esc).join(',')).join('\r\n');const blob=new Blob([csv],{type:'text/csv;charset=utf-8'});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=name.replace(/\.xlsx$/i,'.csv');document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1500);}
function exportTemplate(rows,name,sheet,widths){try{if(!window.XLSX)throw new Error('SheetJS chưa tải');const ws=XLSX.utils.aoa_to_sheet(rows);ws['!cols']=widths.map(w=>({wch:w}));ws['!autofilter']={ref:ws['!ref']};const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,sheet);XLSX.writeFile(wb,name,{compression:true});}catch(e){fallbackCSV(rows,name);alert('Trình duyệt đã tải bản CSV dự phòng. Bạn có thể mở và chỉnh sửa file này bằng Excel.');}}
window.downloadPO=function(){const rows=[['PO','Mã CH','Tên CH','Mã hàng','Tên hàng','PNK','CHECK']];TEMPLATE_STORES.forEach(([sap,name],i)=>rows.push(['PO-'+String(i+1).padStart(3,'0'),sap,name,'','','','CHƯA NHẬP']));exportTemplate(rows,'Mau_Check_PO_18_CH.xlsx','CHECK',[16,10,32,14,32,16,16]);};
window.downloadSchedule=function(){
  const rows=[['Mã nhân viên','Mã infotype','Mã vùng','Ngày bắt đầu','Ngày kết thúc','Mã ca làm việc','Giờ vào','Giờ ra']];
  Object.entries(OPS_CHT).forEach(([mnv])=>rows.push([mnv,'','','','','','','']));
  exportTemplate(rows,'Mau_Du_Lieu_Lich_Tu_OPS.xlsx','DỮ LIỆU OPS',[16,14,12,16,16,16,12,12]);
};
window.parseSchedule=function(workbook){
  const completed=new Set();let matchedRows=0;
  workbook.SheetNames.forEach(sheetName=>{
    const data=XLSX.utils.sheet_to_json(workbook.Sheets[sheetName],{header:1,defval:'',raw:false});
    let header=-1,map={};
    for(let i=0;i<Math.min(data.length,20);i++){
      const normalized=data[i].map(plain);
      const mnvIndex=normalized.findIndex(x=>x==='MA NHAN VIEN'||x.includes('MA NHAN VIEN'));
      const shiftIndex=normalized.findIndex(x=>x==='MA CA LAM VIEC'||x.includes('MA CA'));
      if(mnvIndex>=0&&shiftIndex>=0){header=i;map.mnv=mnvIndex;map.shift=shiftIndex;map.start=normalized.findIndex(x=>x.includes('NGAY BAT DAU'));map.end=normalized.findIndex(x=>x.includes('NGAY KET THUC'));break;}
    }
    if(header<0)return;
    for(let i=header+1;i<data.length;i++){
      const row=data[i],mnv=cleanMNV(row[map.mnv]),sapCode=OPS_CHT[mnv],shift=String(row[map.shift]??'').trim();
      if(!sapCode)continue;
      matchedRows++;
      if(shift)completed.add(sapCode);
    }
  });
  if(!matchedRows){alert('Không tìm thấy dữ liệu OPS đúng mẫu. File cần có cột Mã nhân viên và Mã ca làm việc.');return;}
  state.scheduleDone=SAPS.filter(s=>completed.has(s));
  save();
};

document.addEventListener('DOMContentLoaded',()=>{
  const sample=document.getElementById('sample');if(sample)sample.remove();
  const parse=document.getElementById('parse');if(parse){parse.style.width='100%';parse.style.flex='1 1 100%';}
  const reportName=document.getElementById('rn'),reportInput=document.getElementById('ri');
  if(reportName){if(reportName.value.trim()==='Báo cáo Kho-HCXL')reportName.value='';reportName.placeholder='Nhập tên báo cáo';}
  if(reportInput){if(reportInput.value.trim().startsWith('Báo cáo Kho-HCXL'))reportInput.value='';reportInput.placeholder='Dán danh sách mã cửa hàng đã hoàn thành vào đây';}
  const schedulePanel=document.getElementById('schedule');
  if(schedulePanel){const help=schedulePanel.querySelector('.help');if(help)help.textContent='Upload trực tiếp file xuất từ OPS có các cột: Mã nhân viên, Mã infotype, Mã vùng, Ngày bắt đầu, Ngày kết thúc, Mã ca làm việc, Giờ vào, Giờ ra.';const button=schedulePanel.querySelector('.template .btn');if(button)button.textContent='📥 Tải file mẫu dữ liệu OPS';}
  try{const saved=JSON.parse(localStorage.getItem('leader_ops_v3')||'{}');if(saved.reportName==='Báo cáo Kho-HCXL'){saved.reportName='';saved.reportDone=[];localStorage.setItem('leader_ops_v3',JSON.stringify(saved));}}catch(e){}
  const panel=document.getElementById('reminder'),msg=document.getElementById('msg'),copy=document.getElementById('copy');if(!panel||!msg||!copy)return;
  const form=document.createElement('div');form.className='card';form.innerHTML='<h2>Tên nhắc nhở trên Zalo</h2><p class="help">Nhập tiêu đề muốn hiển thị ở đầu nội dung gửi nhóm.</p><input id="zaloTitle" value="">';panel.insertBefore(form,panel.firstChild);
  const status=document.createElement('div');status.id='copyStatus';status.setAttribute('role','status');status.setAttribute('aria-live','polite');status.style.cssText='display:none;margin:12px 0 0;padding:12px 14px;border-radius:10px;background:#dcfce7;color:#166534;font-weight:800;text-align:center;border:1px solid #86efac';copy.parentElement.parentElement.appendChild(status);
  const toast=document.createElement('div');toast.style.cssText='display:none;position:fixed;right:22px;bottom:22px;z-index:9999;background:#166534;color:white;padding:13px 18px;border-radius:10px;font-weight:800;box-shadow:0 8px 28px rgba(0,0,0,.25)';toast.textContent='✅ Sao chép thành công!';document.body.appendChild(toast);
  const title=document.getElementById('zaloTitle');title.value=localStorage.getItem('leader_ops_zalo_title')||'NHẮC HOÀN THÀNH CÔNG VIỆC';
  function applyTitle(){const current=msg.textContent||'',lines=current.split('\n');if(lines.length){lines[0]='📢 '+(title.value.trim()||'NHẮC HOÀN THÀNH CÔNG VIỆC');const next=lines.join('\n');if(next!==current)msg.textContent=next;}}
  title.addEventListener('input',()=>{localStorage.setItem('leader_ops_zalo_title',title.value);applyTitle();status.style.display='none';});new MutationObserver(applyTitle).observe(msg,{childList:true,characterData:true,subtree:true});applyTitle();
  copy.onclick=async()=>{applyTitle();try{await navigator.clipboard.writeText(msg.textContent);status.textContent='✅ ĐÃ SAO CHÉP THÀNH CÔNG. Bạn có thể dán nội dung vào Zalo.';status.style.display='block';toast.style.display='block';copy.textContent='✅ Đã sao chép';setTimeout(()=>{copy.textContent='Sao chép tin nhắn';toast.style.display='none';},3000);}catch(e){status.textContent='⚠️ Không thể sao chép tự động. Hãy chọn và sao chép nội dung bên dưới.';status.style.background='#fff7ed';status.style.color='#9a3412';status.style.borderColor='#fdba74';status.style.display='block';}};
});