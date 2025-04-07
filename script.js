function calculateMoney() {
    const rows = document.querySelectorAll("#currencyTable tbody tr");
    let total = 0;

    rows.forEach(row => {
        const quantity = parseFloat(row.querySelector(".quantity").value) || 0;
        const rate = parseFloat(row.querySelector(".rate").value) || 0;
        const value = (quantity * rate * 1000);
        row.querySelector(".value").innerText = value.toLocaleString();
        total += value;
    });

    document.getElementById("totalValue").innerText = total.toLocaleString();
}

function saveToExcel() {
    const rows = document.querySelectorAll("#currencyTable tbody tr");
    let csvContent = "Loại Tiền,Số Lượng,Tỷ Giá,Giá Trị\n";

    rows.forEach(row => {
        const currency = row.cells[0].innerText;
        const quantity = row.querySelector(".quantity").value || "0";
        const rate = row.querySelector(".rate").value || "0";
        const value = row.querySelector(".value").innerText || "0";
        csvContent += `${currency},${quantity},${rate},${value}\n`;
    });

    csvContent += `Tổng số tiền,,,${document.getElementById("totalValue").innerText}\n`;

    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    link.setAttribute("href", url);
    link.setAttribute("download", "currency_data.csv");
    link.click();
}