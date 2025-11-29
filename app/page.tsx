"use client";

import { saveAs } from "file-saver";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableCell,
  TableRow,
  AlignmentType,
  WidthType,
  BorderStyle,
} from "docx";

export default function ExportDocx() {
  const exportDoc = async () => {
    const doc = new Document({
      styles: {
        default: {
          document: {
            run: {
              font: "Times New Roman", // <====== FONT MẶC ĐỊNH
              size: 24, // 12pt (mặc định)
            },
          },
        },
      },
      sections: [
        {
          properties: {
            page: {
              size: {
                orientation: "portrait", // hoặc "landscape"
                width: 11906, // A4 ngang: 16838, dọc: 11906 (twip)
                height: 16838, // A4 dọc: 16838
              },
              margin: {
                top: 1440, // 1 inch = 1440 twip
                right: 720, // 0.5 inch
                bottom: 1440,
                left: 720,
              },
            },
          },
          children: [
            // 🇻🇳 Quốc hiệu - tiêu ngữ
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: "(Ban hành kèm theo Thông tư số 34/2017/TT-BGTVT ngày 06 tháng 9 năm 2019 của Bộ trưởng bộ Giao thông vận tải)",
                  size: 20, // cỡ chữ = 9pt ( = 9 * 2 )
                  italics: true,
                }),
              ],
            }),
            new Paragraph(""),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM",
                  bold: true,
                  size: 28,
                }),
              ],
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: "Độc lập - Tự do - Hạnh phúc",
                  bold: true,
                  size: 28,
                }),
              ],
            }),

            new Paragraph(""),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: "---------------",
                  bold: true,
                }),
              ],
            }),
            new Paragraph(""),
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [
                new TextRun({
                  text: "Hạ Long, ngày…… tháng…… năm 20…",
                }),
              ],
            }),
            new Paragraph(""), // dòng trống

            // 🔴 Tiêu đề chính màu đỏ
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: "DANH SÁCH",
                  bold: true,
                }),
              ],
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: "HÀNH KHÁCH VẬN TẢI ĐƯỜNG THỦY NỘI ĐỊA",
                  bold: true,
                }),
              ],
            }),
            new Paragraph(""),

            // Thông tin tàu
            new Paragraph({
              children: [
                new TextRun({ text: "Tên phương tiện: " }),
                new TextRun({ text: "ABC123" }),
                new TextRun({ text: " Số đăng ký: " }),
                new TextRun("QN-9999"),
                new TextRun({ text: " Sức chở: " }),
                new TextRun("10 người."),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "Tên chủ phương tiện: " }),
                new TextRun("{CHU_TAU.ten_chu_tau}"),
              ],
            }),

            new Paragraph({
              children: [
                new TextRun({ text: "Địa chỉ: " }),
                new TextRun("Hà Nội"),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "Tên thuyền trưởng: " }),
                new TextRun("{THUYEN_TRUONG.tt_so_giay_phep_lai_tau}"),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "SĐT:" }),
                new TextRun("{THUYEN_TRUONG.tt_so_giay_phep_lai_tau}"),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "Số lượng thuyền viên:" }),
                new TextRun("{THUYEN_TRUONG.tt_so_giay_phep_lai_tau}"),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "Số lượng phục vụ:" }),
                new TextRun("{THUYEN_TRUONG.tt_so_giay_phep_lai_tau}"),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "Tuyến vận tải: " }),
                new TextRun("......................................."),
                new TextRun({ text: "Hành trình VHL:" }),
                new TextRun("{ HANH_TRINH.ten_hanh_trinh}"),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "Thời gian rời bến: hồi hour(BOOKINGS.tt_ngay_di) giờ minute(BOOKINGS.tt_ngay_di), ngày BOOKINGS.tt_ngay_di",
                }),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "Tổng khách: BOOKINGS.th_tong_so_khach hoặc tính SUM(so_luong) Quốc tịch: Việt Nam BOOKINGS.th_tong_khach_viet_nam  người; nước ngoài BOOKINGS.th_tong_khach_nuoc_ngoai người",
                }),
              ],
            }),
            new Paragraph(""),

            // 📋 Bảng hành khách
            new Table({
              width: { size: 11906 - 720 - 720, type: WidthType.DXA },
              rows: [
                new TableRow({
                  children: [
                    cell("STT", true, 1000), // ~0.7 inch
                    cell("Họ và tên", true, 4000), // ~2.8 inch
                    cell("Năm sinh (tuổi)", true, 1500),
                    cell("Nam/Nữ", true, 1500),
                    cell("Quốc tịch", true, 2000),
                    cell("Ghi chú", true, 2000),
                  ],
                }),
                new TableRow({
                  children: [
                    cell("1"),
                    cell("Nguyễn Văn A"),
                    cell("1990"),
                    cell("Nam"),
                    cell("Việt Nam"),
                    cell(""),
                  ],
                }),
                new TableRow({
                  children: [
                    cell("2"),
                    cell("Trần Thị B"),
                    cell("1992"),
                    cell("Nữ"),
                    cell("Việt Nam"),
                    cell(""),
                  ],
                }),
              ],
            }),

            new Paragraph(""),

            // Footer ký tên
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: "Tổng số hành khách BOOKINGS.th_tong_so_khach hoặc tính SUM(so_luong) người (bằng chữ {dùng hàm chuyển số sang chữ} người)",
                }),
              ],
            }),
            new Paragraph(""),
            new Table({
              width: {
                size: 100,
                type: WidthType.PERCENTAGE,
              },
              borders: {
                top: { style: BorderStyle.SINGLE, size: 1, color: "FFFFFF" }, // trắng
                bottom: { style: BorderStyle.SINGLE, size: 1, color: "FFFFFF" },
                left: { style: BorderStyle.SINGLE, size: 1, color: "FFFFFF" },
                right: { style: BorderStyle.SINGLE, size: 1, color: "FFFFFF" },
                insideHorizontal: {
                  style: BorderStyle.SINGLE,
                  size: 1,
                  color: "FFFFFF",
                },
                insideVertical: {
                  style: BorderStyle.SINGLE,
                  size: 1,
                  color: "FFFFFF",
                },
              },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      children: [
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: "ĐẠI DIỆN ĐƠN VỊ KHAI THÁC CẢNG, BẾN",
                              bold: true,
                            }),
                          ],
                          alignment: AlignmentType.CENTER,
                        }),
                      ],
                    }),
                    new TableCell({
                      children: [
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: "NGƯỜI LẬP DANH SÁCH",
                              bold: true,
                            }),
                          ],
                          alignment: AlignmentType.CENTER,
                        }),
                      ],
                    }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({
                      children: [
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: "(ký, ghi rõ họ, tên)",
                            }),
                          ],
                          alignment: AlignmentType.CENTER,
                        }),
                      ],
                    }),
                    new TableCell({
                      children: [
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: "(ký, ghi rõ họ, tên)",
                            }),
                          ],
                          alignment: AlignmentType.CENTER,
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          ],
        },
      ],
    });

    // cell helper
    function cell(text: string, header = false, width?: number) {
      return new TableCell({
        width: width
          ? { size: width, type: WidthType.PERCENTAGE } // dùng %
          : undefined,
        verticalAlign: "center",
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text, bold: header })],
          }),
        ],
      });
    }
    const blob = await Packer.toBlob(doc);
    saveAs(blob, "danh_sach_hanh_khach.docx");
  };

  return (
    <button
      onClick={exportDoc}
      style={{ padding: "10px 20px", background: "green", color: "#fff" }}
    >
      Xuất DOCX đẹp
    </button>
  );
}
