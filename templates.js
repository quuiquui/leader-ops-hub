const TEMPLATE_STORES=[['3888','53 Hoàng Bá Bích'],['4044','389 Đường N6'],['4112','38 Đặng Văn Trơn'],['4139','157-159 Phan Đình Phùng'],['4163','3.9 Nguyễn Văn Tỏ'],['4227','869 Hoàng Tam Kỳ'],['4324','A83 đường 5'],['4354','81 Nguyễn Hoàng'],['4465','G1 Khu 94 Long Đức 1'],['4468','152 Đinh Quang Ân'],['4673','27 Phùng Hưng'],['5140','175-177 đường N16'],['2DEL','Nguyễn Hoàng'],['6055','G1/31 Tổ 19'],['6138','Tổ 13 KP Vườn Dừa'],['6390','167 Ngô Quyền'],['6488','1111 Bùi Văn Hòa'],['2BH7','Lô 17-18 KDC Bình Dương']];
function fallbackCSV(rows,name){const esc=v=>'"'+String(v??'').replace(/"/g,'""')+'"';const csv='\ufeff'+rows.map(r=>r.map(esc).join(',')).join('\r\n');const blob=new Blob([csv],{type:'text/csv;charset=utf-8'});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=name.replace(/\.xlsx$/i,'.csv');document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1500);}
function exportTemplate(rows,name,sheet,widths){try{if(!window.XLSX)throw new Error('SheetJS chưa tải');const ws=XLSX.utils.aoa_to_sheet(rows);ws['!cols']=widths.map(w=>({wch:w}));ws['!autofilter']={ref:ws['!ref']};const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,sheet);XLSX.writeFile(wb,name,{compression:true});}catch(e){fallbackCSV(rows,name);alert('Trình duyệt đã tải bản CSV dự phòng. Bạn có thể mở và chỉnh sửa file này bằng Excel.');}}
window.downloadPO=function(){const rows=[['PO','Mã CH','Tên CH','Mã hàng','Tên hàng','PNK','CHECK']];TEMPLATE_STORES.forEach(([sap,name],i)=>rows.push(['PO-'+String(i+1).padStart(3,'0'),sap,name,'','','','CHƯA NHẬP']));exportTemplate(rows,'Mau_Check_PO_18_CH.xlsx','CHECK',[16,10,32,14,32,16,16]);};
window.downloadSchedule=function(){const rows=[['Mã SAP','Tên cửa hàng','MNV','CHT/CHP đứng chính','Thứ 2','Thứ 3','Thứ 4','Thứ 5','Thứ 6','Thứ 7','Chủ nhật']];TEMPLATE_STORES.forEach(([sap,name])=>rows.push([sap,name,'','','','','','','','','']));exportTemplate(rows,'Mau_Lich_Lam_Viec_18_CH.xlsx','LỊCH LÀM VIỆC',[10,32,14,24,12,12,12,12,12,12,12]);};

document.addEventListener('DOMContentLoaded',()=>{
  const sample=document.getElementById('sample');
  if(sample)sample.remove();
  const parse=document.getElementById('parse');
  if(parse){parse.style.width='100%';parse.style.flex='1 1 100%';}

  const panel=document.getElementById('reminder');
  const msg=document.getElementById('msg');
  const copy=document.getElementById('copy');
  if(!panel||!msg||!copy)return;

  const form=document.createElement('div');
  form.className='card';
  form.innerHTML='<h2>Tên nhắc nhở trên Zalo</h2><p class="help">Nhập tiêu đề muốn hiển thị ở đầu nội dung gửi nhóm.</p><input id="zaloTitle" value="">';
  panel.insertBefore(form,panel.firstChild);

  const status=document.createElement('div');
  status.id='copyStatus';
  status.setAttribute('role','status');
  status.setAttribute('aria-live','polite');
  status.style.cssText='display:none;margin:12px 0 0;padding:12px 14px;border-radius:10px;background:#dcfce7;color:#166534;font-weight:800;text-align:center;border:1px solid #86efac';
  copy.parentElement.parentElement.appendChild(status);

  const toast=document.createElement('div');
  toast.id='copyToast';
  toast.style.cssText='display:none;position:fixed;right:22px;bottom:22px;z-index:9999;background:#166534;color:white;padding:13px 18px;border-radius:10px;font-weight:800;box-shadow:0 8px 28px rgba(0,0,0,.25)';
  toast.textContent='✅ Sao chép thành công!';
  document.body.appendChild(toast);

  const title=document.getElementById('zaloTitle');
  title.value=localStorage.getItem('leader_ops_zalo_title')||'NHẮC HOÀN THÀNH CÔNG VIỆC';
  function applyTitle(){const current=msg.textContent||'';const lines=current.split('\n');if(lines.length){lines[0]='📢 '+(title.value.trim()||'NHẮC HOÀN THÀNH CÔNG VIỆC');const next=lines.join('\n');if(next!==current)msg.textContent=next;}}
  title.addEventListener('input',()=>{localStorage.setItem('leader_ops_zalo_title',title.value);applyTitle();status.style.display='none';});
  new MutationObserver(applyTitle).observe(msg,{childList:true,characterData:true,subtree:true});
  applyTitle();

  copy.onclick=async()=>{
    applyTitle();
    try{
      await navigator.clipboard.writeText(msg.textContent);
      status.textContent='✅ ĐÃ SAO CHÉP THÀNH CÔNG. Bạn có thể dán nội dung vào Zalo.';
      status.style.display='block';
      toast.style.display='block';
      copy.textContent='✅ Đã sao chép';
      setTimeout(()=>{copy.textContent='Sao chép tin nhắn';toast.style.display='none';},3000);
    }catch(e){
      status.textContent='⚠️ Không thể sao chép tự động. Hãy chọn và sao chép nội dung bên dưới.';
      status.style.background='#fff7ed';status.style.color='#9a3412';status.style.borderColor='#fdba74';status.style.display='block';
    }
  };
});