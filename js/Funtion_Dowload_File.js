// Ngon ngu tieng viet
const chung_Chi_Rua_Toi_Va_Them_Suc = '/File_word_tanhung/VietNamese/CHỨNG CHỈ RỬA TỘI VÀ THÊM SỨC.docx';
const chung_Chi_Rua_Toi = '/File_word_tanhung/VietNamese/CHỨNG CHỈ RỬA TỘI.docx';
const chung_Chi_Them_Suc = '/File_word_tanhung/VietNamese/CHỨNG CHỈ THÊM SỨC.docx';
const chung_Thu_Hon_Phoi = '/File_word_tanhung/VietNamese/CHỨNG THƯ HÔN PHỐI.docx';
const don_Xin_Nhap_Xu = '/File_word_tanhung/VietNamese/ĐƠN XIN NHẬP XỨ.docx';
const don_Xin_Phep_Chuan_Hon_Nhan_Khac_Ton_Giao = '/File_word_tanhung/VietNamese/ĐƠN XIN PHÉP CHUẨN HÔN NHÂN KHÁC TÔN GIÁO.docx';
const don_Xin_Rua_Toi_Tre_Em = '/File_word_tanhung/VietNamese/ĐƠN XIN RỬA TỘI TRẺ EM.docx';
const don_Xin_Thua_Tac_Vien_Ngoai_Le_Trao_Minh_Thanh_Chua = '/File_word_tanhung/VietNamese/ĐƠN XIN THỪA TÁC VIÊN NGOẠI LỆ TRAO MÌNH THÁNH CHÚA.docx';
const don_Xin_Chao_Thua_Tac_Vu_Ngoai_Le_Trao_Minh_Thanh_Chua = '/File_word_tanhung/VietNamese/ĐƠN XIN TRAO THỪA TÁC VỤ NGOẠI LỆ TRAO MÌNH THÁNH CHÚA.docx';
const gioi_Thieu_Hon_Phoi = '/File_word_tanhung/VietNamese/GIỚI THIỆU HÔN PHỐI.docx';
const ket_Qua_Rao_Hon_Phoi = '/File_word_tanhung/VietNamese/KẾT QUẢ RAO HÔN PHỐI.docx';
const to_Khai_Hon_Phoi_Ben_Cong_Giao = '/File_word_tanhung/VietNamese/TỜ KHAI HÔN PHỐI (Phép Chuẩn - bên Công Giáo tự tay viết).docx';
const to_Khai_Hon_Phoi_Ngoai_Cong_Giao = '/File_word_tanhung/VietNamese/TỜ KHAI HÔN PHỐI(Phép Chuẩn - bên ngoài Công Giáo tự tay viết).docx';
const to_Rao_Hon_Phoi = '/File_word_tanhung/VietNamese/TỜ RAO HÔN PHỐI.docx';
const xac_Nhan_Gioi_Thieu_Hon_Phoi = '/File_word_tanhung/VietNamese/XÁC NHẬN và GIỚI THIỆU HÔN PHỐI.docx';

// Ngonn ngu tieng anh
const CERTIFICATE_OF_BAPTISM_AND_CONFIRMATION = '/File_word_tanhung/English/CERTIFICATE OF BAPTISM AND CONFIRMATION.docx';
const CERTIFICATE_OF_BAPTISM = '/File_word_tanhung/English/CERTIFICATE OF BAPTISM.docx';
const CERTIFICATE_OF_CONFIRMATION = '/File_word_tanhung/English/CERTIFICATE OF CONFIRMATION.docx';
const CERTIFICATE_OF_MARRIAGE = '/File_word_tanhung/English/CERTIFICATE OF MARRIAGE.docx';
const NOTIFICATION_OF_MARRIAGE = '/File_word_tanhung/English/NOTIFICATION OF MARRIAGE.docx';

// Ngon ngu tieng phap
const CERTIFICAT_DE_BAPTEME_et_DE_CONFIRMATION = '/File_word_tanhung/French/CERTIFICAT DE BAPTÊME et DE CONFIRMATION.docx';
const CERTIFICAT_DE_BAPTEME = '/File_word_tanhung/French/CERTIFICAT DE BAPTÊME.docx';
const CERTIFICAT_DE_CONFIRMATION = '/File_word_tanhung/French/CERTIFICAT DE CONFIRMATION.docx';
const CERTIFICAT_DE_MARRIAGE = '/File_word_tanhung/French/CERTIFICAT DE MARRIAGE.docx';
const NOTIFICATION_DE_MARIAGE = '/File_word_tanhung/French/NOTIFICATION DE MARIAGE.docx';


// Phuong thuc tai file tieng viet
function download_Chung_Chi_Rua_Toi_Va_Them_Suc() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = chung_Chi_Rua_Toi_Va_Them_Suc; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'CHỨNG CHỈ RỬA TỘI VÀ THÊM SỨC.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

function download_Chung_Chi_Rua_Toi() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = chung_Chi_Rua_Toi; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'CHỨNG CHỈ RỬA TỘI.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

function download_Chung_Chi_Them_Suc() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = chung_Chi_Them_Suc; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'CHỨNG CHỈ THÊM SỨC.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

function download_Chung_Thu_Hon_Phoi() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = chung_Thu_Hon_Phoi; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'CHỨNG THƯ HÔN PHỐI.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

function download_Don_Xin_Nhap_Xu() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = don_Xin_Nhap_Xu; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'ĐƠN XIN NHẬP XỨ.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

function download_Don_Xin_Phep_Chuan_Hon_Nhan_Khac_Ton_Giao() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = don_Xin_Phep_Chuan_Hon_Nhan_Khac_Ton_Giao; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'ĐƠN XIN PHÉP CHUẨN HÔN NHÂN KHÁC TÔN GIÁO.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

function download_Don_Xin_Rua_Toi_Tre_Em() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = don_Xin_Rua_Toi_Tre_Em; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'ĐƠN XIN RỬA TỘI TRẺ EM.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

function download_Don_Xin_Thua_Tac_Vien_Ngoai_Le_Trao_Minh_Thanh_Chua() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = don_Xin_Thua_Tac_Vien_Ngoai_Le_Trao_Minh_Thanh_Chua; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'ĐƠN XIN THỪA TÁC VIÊN NGOẠI LỆ TRAO MÌNH THÁNH CHÚA.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

function download_Don_Xin_Chao_Thua_Tac_Vu_Ngoai_Le_Trao_Minh_Thanh_Chua() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = don_Xin_Chao_Thua_Tac_Vu_Ngoai_Le_Trao_Minh_Thanh_Chua; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'ĐƠN XIN TRAO THỪA TÁC VỤ NGOẠI LỆ TRAO MÌNH THÁNH CHÚA.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

function download_Gioi_Thieu_Hon_Phoi() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = gioi_Thieu_Hon_Phoi; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'GIỚI THIỆU HÔN PHỐI.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

function download_Ket_Qua_Rao_Hon_Phoi() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = ket_Qua_Rao_Hon_Phoi; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'KẾT QUẢ RAO HÔN PHỐI.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

function download_To_Khai_Hon_Phoi_Ben_Cong_Giao() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = to_Khai_Hon_Phoi_Ben_Cong_Giao; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'TỜ KHAI HÔN PHỐI (Phép Chuẩn - bên Công Giáo tự tay viết).docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

function download_To_Khai_Hon_Phoi_Ngoai_Cong_Giao() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = to_Khai_Hon_Phoi_Ngoai_Cong_Giao; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'TỜ KHAI HÔN PHỐI(Phép Chuẩn - bên ngoài Công Giáo tự tay viết).docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

function download_To_Rao_Hon_Phoi() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = to_Rao_Hon_Phoi; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'TỜ RAO HÔN PHỐI.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

function download_Xac_Nhan_Gioi_Thieu_Hon_Phoi() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = xac_Nhan_Gioi_Thieu_Hon_Phoi; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'XÁC NHẬN và GIỚI THIỆU HÔN PHỐI.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

//Phuong thuc tai file tieng anh
function download_CERTIFICATE_OF_BAPTISM_AND_CONFIRMATION() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = CERTIFICATE_OF_BAPTISM_AND_CONFIRMATION; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'CERTIFICATE OF BAPTISM AND CONFIRMATION.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

function download_CERTIFICATE_OF_BAPTISM() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = CERTIFICATE_OF_BAPTISM; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'CERTIFICATE OF BAPTISM.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

function download_CERTIFICATE_OF_CONFIRMATION() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = CERTIFICATE_OF_CONFIRMATION; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'CERTIFICATE OF CONFIRMATION.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

function download_CERTIFICATE_OF_MARRIAGE() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = CERTIFICATE_OF_MARRIAGE; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'CERTIFICATE OF MARRIAGE.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

function download_NOTIFICATION_OF_MARRIAGE() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = NOTIFICATION_OF_MARRIAGE; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'NOTIFICATION OF MARRIAGE.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

//Phuong thuc tai file tieng phap
function download_CERTIFICAT_DE_BAPTEME_et_DE_CONFIRMATION() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = CERTIFICAT_DE_BAPTEME_et_DE_CONFIRMATION; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'CERTIFICAT DE BAPTÊME et DE CONFIRMATION.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

function download_CERTIFICAT_DE_BAPTEME() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = CERTIFICAT_DE_BAPTEME; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'CERTIFICAT DE BAPTÊME.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

function download_CERTIFICAT_DE_CONFIRMATION() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = CERTIFICAT_DE_CONFIRMATION; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'CERTIFICAT DE CONFIRMATION.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

function download_CERTIFICAT_DE_MARRIAGE() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = CERTIFICAT_DE_MARRIAGE; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'CERTIFICAT DE MARRIAGE.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

function download_NOTIFICATION_DE_MARIAGE() {
    //thư viên Swal hỗ trợ
    let timerInterval;
    Swal.fire({
        title: "Tải file từ hệ thống",
        html: "Vui lòng đợi hệ thống xử lý",
        timer: 1800,
        timerProgressBar: true,
        didOpen: () => {
            Swal.showLoading();
            const timer = Swal.getPopup().querySelector("b");
            timerInterval = setInterval(() => {
                timer.textContent = `${Swal.getTimerLeft()}`;
            }, 100);
        },
        willClose: () => {
            clearInterval(timerInterval);
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.timer) {
            //lấy file excel trả về cho người dùng 
            const link = document.createElement('a');
            link.href = NOTIFICATION_DE_MARIAGE; // Đường dẫn đến file Excel trên máy chủ
            link.download = 'NOTIFICATION DE MARIAGE.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });
}

