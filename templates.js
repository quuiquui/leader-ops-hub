document.addEventListener("DOMContentLoaded", function () {
  // Đổi tên website
  document.title = "Leader Ops Hub";

  const heading = document.querySelector("header h1");
  if (heading) {
    heading.textContent = "LEADER OPS HUB";
  }

  // Xóa nút Điền mẫu
  const sampleButton = document.getElementById("sample");
  if (sampleButton) {
    sampleButton.remove();
  }

  // Xóa chữ mẫu trong ô Tên báo cáo
  const reportName = document.getElementById("rn");
  if (reportName) {
    reportName.value = "";
    reportName.placeholder = "Nhập tên báo cáo";
  }

  // Xóa danh sách mẫu trong khung nhập
  const reportInput = document.getElementById("ri");
  if (reportInput) {
    reportInput.value = "";
    reportInput.placeholder =
      "Dán danh sách mã cửa hàng đã hoàn thành vào đây";
  }

  // Xóa dữ liệu mẫu cũ đã lưu trên trình duyệt
  try {
    const saved = JSON.parse(
      localStorage.getItem("leader_ops_v3") || "{}"
    );

    if (saved.reportName === "Báo cáo Kho-HCXL") {
      saved.reportName = "";
      saved.reportDone = [];
      localStorage.setItem(
        "leader_ops_v3",
        JSON.stringify(saved)
      );
    }
  } catch (error) {
    console.log("Không có dữ liệu mẫu cũ cần xóa.");
  }

  console.log("Leader Ops Hub đã được cập nhật.");
});
