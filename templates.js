document.addEventListener('DOMContentLoaded',()=>{
  const stores=[
    ['3888','53 Hoàng Bá Bích'],['4044','389 Đường N6'],['4112','38 Đặng Văn Trơn'],['4139','157-159 Phan Đình Phùng'],['4163','3.9 Nguyễn Văn Tỏ'],['4227','869 Hoàng Tam Kỳ'],['4324','A83 đường 5'],['4354','81 Nguyễn Hoàng'],['4465','G1 Khu 94 Long Đức 1'],['4468','152 Đinh Quang Ân'],['4673','27 Phùng Hưng'],['5140','175-177 đường N16'],['2DEL','Nguyễn Hoàng'],['6055','G1/31 Tổ 19'],['6138','Tổ 13 KP Vườn Dừa'],['6390','167 Ngô Quyền'],['6488','1111 Bùi Văn Hòa'],['2BH7','Lô 17-18 KDC Bình Dương']
  ];
  function style(ws,widths){
    ws['!cols']=widths.map(w=>({wch:w}));
    ws['!freeze']={xSplit:0,ySplit:1};
    ws['!autofilter']={ref:ws['!ref']};
  }
  function downloadPO(){
    const rows=[['PO','Mã CH','Tên CH','Mã hàng','Tên hàng','PNK','CHECK']];
    stores.forEach(([sap,name],i)=>rows.push(['PO-'+String(i+1).padStart(3,'0'),sap,name,'','','','CHƯA NHẬP']));
    const ws=XLSX.utils.aoa_to_sheet(rows); style(ws,[16,10,32,14,32,16,16]);
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'CHECK');
    XLSX.writeFile(wb,'Mau_Check_PO_18_CH.xlsx');
  }
  function downloadSchedule(){
    const rows=[['Mã SAP','Tên cửa hàng','MNV','CHT/CHP đứng chính','Thứ 2','Thứ 3','Thứ 4','Thứ 5','Thứ 6','Thứ 7','Chủ nhật']];
    stores.forEach(([sap,name])=>rows.push([sap,name,'','','','','','','','','']));
    const ws=XLSX.utils.aoa_to_sheet(rows); style(ws,[10,32,14,24,12,12,12,12,12,12,12]);
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'LỊCH LÀM VIỆC');
    XLSX.writeFile(wb,'Mau_Lich_Lam_Viec_18_CH.xlsx');
  }
  function addButton(inputId,label,handler){
    const input=document.getElementById(inputId); if(!input)return;
    const btn=document.createElement('button'); btn.type='button'; btn.className='btn light'; btn.textContent=label; btn.style.margin='10px 6px 0'; btn.onclick=handler;
    input.parentElement.appendChild(document.createElement('br')); input.parentElement.appendChild(btn);
  }
  addButton('pf','📥 Tải file Excel mẫu Check PO',downloadPO);
  addButton('sf','📥 Tải file Excel mẫu Lịch làm việc',downloadSchedule);
});