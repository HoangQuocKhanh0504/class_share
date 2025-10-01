const express = require("express");
const multer = require("multer");
const AdmZip = require("adm-zip");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
const cors = require("cors");

const app = express();
const PORT = 3000;

app.use(cors());
app.use(express.static("public"));
app.use("/output", express.static("output"));

const upload = multer({ dest: "uploads/" });

let danhSach = []; // Lưu thông tin học sinh + file

// Hàm loại bỏ dấu và khoảng trắng
const removeVietnameseTones = (str) => {
  str = str.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  str = str.replace(/\s+/g, "");
  return str;
};

// Upload danh sách Excel
app.post("/upload-list", upload.single("file"), (req, res) => {
  try {
    const workbook = XLSX.readFile(req.file.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);

    danhSach = data.map((row, i) => {
      const hoten = row["Họ tên"] || row["Ho ten"] || "Không rõ";
      const hotenNoSign = removeVietnameseTones(hoten); // dùng cho filename
      return {
        stt: i + 1,
        hoten,           // Hiển thị trên index
        hotenNoSign,     // Dùng làm tên file
        file: null,
        size: null,
        link: null,
      };
    });

    fs.unlinkSync(req.file.path);
    res.json({ success: true, danhSach });
  } catch (err) {
    res.json({ success: false, msg: err.message });
  }
});

// Upload project ZIP
app.post("/upload-project", upload.single("file"), (req, res) => {
  if (!danhSach.length) return res.json({ success: false, msg: "Chưa có danh sách" });

  try {
    const zip = new AdmZip(req.file.path);
    const extractPath = path.join(__dirname, "uploads", "project");
    if (fs.existsSync(extractPath)) fs.rmSync(extractPath, { recursive: true });
    zip.extractAllTo(extractPath, true);

    // Tạo file ZIP riêng cho từng học sinh
    danhSach.forEach(student => {
      const zipClone = new AdmZip();

      zip.getEntries().forEach(entry => {
        let content = entry.getData();
        if (!entry.isDirectory && /\.(html|js|css|php|txt|md)$/i.test(entry.entryName)) {
          // Thay 'khanhpc' trong file bằng tên đầy đủ (có dấu)
          let text = content.toString("utf8").replace(/khanhpc/g, student.hoten);
          zipClone.addFile(entry.entryName, Buffer.from(text, "utf8"));
        } else {
          zipClone.addFile(entry.entryName, content);
        }
      });

      const outPath = path.join(__dirname, "output", `${student.hotenNoSign}.zip`);
      zipClone.writeZip(outPath);

      student.file = `${student.hotenNoSign}.zip`;
      student.size = (fs.statSync(outPath).size / 1024).toFixed(2) + " KB";
      student.link = `/output/${student.hotenNoSign}.zip`;
    });

    fs.unlinkSync(req.file.path);
    res.json({ success: true, danhSach });
  } catch (err) {
    res.json({ success: false, msg: err.message });
  }
});

// API danh sách
app.get("/list", (req, res) => {
  res.json(danhSach);
});

app.listen(PORT, () => console.log(`Server chạy tại http://localhost:${PORT}`));
