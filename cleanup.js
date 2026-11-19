document.addEventListener('DOMContentLoaded',()=>{
  const name=document.getElementById('rn');
  const input=document.getElementById('ri');
  let changed=false;
  try{
    if(typeof state!=='undefined' && state.reportName==='Báo cáo Kho-HCXL'){
      state.reportName='';
      state.reportDone=[];
      localStorage.setItem('leader_ops_v3',JSON.stringify(state));
      changed=true;
    }
  }catch(e){}
  if(name && name.value.trim()==='Báo cáo Kho-HCXL'){name.value='';changed=true;}
  if(input && input.value.trim().startsWith('Báo cáo Kho-HCXL')){input.value='';changed=true;}
  if(name)name.placeholder='Nhập tên báo cáo';
  if(input)input.placeholder='Dán danh sách mã cửa hàng đã hoàn thành vào đây';
  if(changed && typeof render==='function'){
    try{render();if(name)name.value='';if(input)input.value='';}catch(e){}
  }
});