/* ─── 設定エリア ─── */
// ★★★ ここにGoogle Apps ScriptのURLを貼り付けてください ★★★
const GAS_URL = 'https://script.google.com/macros/s/AKfycbzMYsFXdslrtYBiFSqBOcXcQRvQ_febxC-68eMqrqCPlaSgrPLuzjdXvtUmtO4t6VyhMw/exec'

const TIMES = ["09:00", "10:00", "11:00", "12:00", "13:00", "14:00", "15:00", "16:00", "17:00", "18:00", "19:00"];
const ADMIN_CAPACITY = 7;
const STUDENT_CAPACITY = 5;

let students = []; // スプレッドシートから読み込む
let reservations = []; // スプレッドシートから読み込む
let currentStudent = null;
let viewDate = new Date();
let adminViewDate = new Date();
let adminListMonth = new Date();
let selectedDateStr = null;
let selectedCell = null;
let examNames = []; 

/**
 * 共通：データのクレンジング
 */
function normalizeDate(input) {
    if (!input) return "";
    let d = new Date(typeof input === 'string' ? input.replace(/-/g, '/') : input);
    if (isNaN(d.getTime())) return "";
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

function normalizeTime(input) {
    if (!input) return "";
    let t = String(input).trim();
    let match = t.match(/(\d{1,2}):(\d{1,2})/);
    if (match) return match[1].padStart(2, '0') + ':' + match[2].padStart(2, '0');
    return t;
}

window.onload = async () => {
    await syncData();
    if (document.getElementById('admin-view')) {
        renderStudentList();
        renderWeeklySchedule();
    }
};

/**
 * 1. データの同期（受講生情報と予約情報の両方を取得）
 */
async function syncData() {
    if(!GAS_URL || GAS_URL.startsWith('ここ')) return;
    const overlay = document.getElementById('loadingOverlay');
    if(overlay) overlay.style.display = 'flex';
    
    try {
        const res = await fetch(GAS_URL + '?t=' + new Date().getTime());
        const data = await res.json();
        
        const cleanId = (id) => String(id).split('.')[0].trim();
        const cleanTime = (t) => {
            const m = String(t).match(/(\d{1,2}:\d{2})/);
            if(!m) return normalizeTime(t);
            return (m[1].length === 4 ? "0" + m[1] : m[1]);
        };

        // 1. 実データの正規化（予約・取消の両方が入ります）
        const rawRes = data.reservations.map(r => ({
            ...r,
            date: normalizeDate(r.date),
            time: cleanTime(r.time),
            studentId: cleanId(r.studentId),
            status: String(r.status || "").trim(),
             pcIndex: (r.pcIndex === undefined || r.pcIndex === null || r.pcIndex === "") ? 0 : Number(r.pcIndex)
        }));

        students = data.students.map(s => ({ ...s, id: cleanId(s.id) }));

        examNames = data.examNames || [];

        console.log("読み込まれた検定名:", examNames); 

        if (document.getElementById('examNameList')) {
            document.getElementById('examNameList').innerText = "登録済み: " + examNames.join(', ');
        }
        // 2. 照合用マップ：ID|日付|時間をキーに「最新のステータス」を保持
        const statusMap = new Map();
        // rawRes.forEach(r => {
        //     statusMap.set(`${r.studentId}|${r.date}|${r.time}`, r.status);
        // });
        rawRes.forEach(r => {
            // キーに pcIndex を含めて、同じ時間の複数予約（検定等）を識別可能にする
            // const key = `${r.studentId}|${r.date}|${r.time}|${r.pcIndex}`;
            const key = r.studentId + "_" + r.date + "_" + r.time + "_" + (r.pcIndex || 0);
            statusMap.set(key, r.status);
            
        });

        // 3. 仮想（固定）予約の生成
        // const statusSet = new Set(rawRes.map(r => `${r.studentId}|${r.date}|${r.time}|${r.pcIndex}`));
        
        const virtual = [];
        const start = new Date(); start.setMonth(start.getMonth() - 1);
        const end = new Date(); end.setMonth(end.getMonth() + 3);

        // ★高速化：日付リストをあらかじめ1回だけ作成しておく
        const dateList = [];
        let dPtr = new Date(start);
        while (dPtr <= end) {
            dateList.push({ dNorm: normalizeDate(dPtr), day: dPtr.getDay() });
            dPtr.setDate(dPtr.getDate() + 1);
        }

        students.forEach(st => {
            let curr = new Date(start);
            if (st.fixedDay !== "" && st.fixedDay !== undefined && st.fixedTime) {
                const targetDay = Number(st.fixedDay);
                const tNorm = cleanTime(st.fixedTime);
                const stId = st.id;

                dateList.forEach(item => {
                    if (item.day === targetDay) {
                        const key = stId + "_" + item.dNorm + "_" + tNorm + "_0";
                        if (!statusMap.has(key)) {
                            virtual.push({
                                date: item.dNorm, time: tNorm, studentId: stId, name: st.name,
                                course: st.course, pcIndex: 0, status: "予約"
                            });
                        }
                    }
                });
                
                // while (curr <= end) {
                //     if (curr.getDay() === targetDay) {
                //         const dNorm = normalizeDate(curr);
                //         // 固定枠はPC1(index 0)を想定
                //         // const key = `${st.id}|${dNorm}|${tNorm}|0`;
                //         const key = st.id + "_" + dNorm + "_" + tNorm + "_0";
                        
                //         // if (!statusMap.has(key)) {
                //         //     virtual.push({
                //         //         date: dNorm, time: tNorm, studentId: st.id, name: st.name,
                //         //         course: st.course, pcIndex: 0, status: "予約"
                //         //     });
                //         // }

                //         if (!statusMap.has(key)) {
                //             virtual.push({
                //                 date: dNorm, 
                //                 time: tNorm, 
                //                 studentId: st.id, 
                //                 name: st.name,
                //                 course: st.course, 
                //                 pcIndex: 0, 
                //                 status: "予約"
                //             });
                //         }

                //     }
                //     curr.setDate(curr.getDate() + 1);
                // }


                // while (curr <= end) {
                //     if (curr.getDay() === targetDay) {
                //         const dNorm = normalizeDate(curr);
                //         const key = `${st.id}|${dNorm}|${tNorm}`;

                //         // ★シートに「予約」または「取消」の記録があれば、固定枠の自動生成をスキップ
                //         if (!statusMap.has(key)) {
                //             virtual.push({
                //                 date: dNorm, time: tNorm, studentId: st.id, name: st.name,
                //                 course: st.course, pcIndex: 0, status: "予約"
                //             });
                //         }
                //     }
                //     curr.setDate(curr.getDate() + 1);
                // }
            }
        });

        // 4. 表示用の最終結合（実データの「予約」 ＋ 生成した固定枠）
        // 「取消」データは、ここで除外されるため画面から消えます
        const realBookings = rawRes.filter(r => r.status === "予約");
        reservations = [...realBookings, ...virtual];

        localStorage.setItem('sch_v5_res', JSON.stringify(reservations));
        localStorage.setItem('sch_v5_students', JSON.stringify(students));

    } catch (e) {
        console.error("Sync Error:", e);
    } finally {
        if(overlay) overlay.style.display = 'none';
        if (document.getElementById('studentList')) renderStudentList();
        if (document.getElementById('weeklyScheduleContainer')) renderWeeklySchedule();
        if (document.getElementById('s_select_list')) updateStudentEditDropdown();
    }
}

/**
 * 2. Googleへのデータ送信関数（予約・取消用）
 */
function sendToGoogleSheet(date, time, status, studentObj, pcIndex) {
    // 予約・取消データは type を指定しない（GAS側のelseに入る）
    const data = { 
        date: normalizeDate(date), 
        time: normalizeTime(time), 
        status: status, 
        studentId: studentObj.id, 
        name: studentObj.name, 
        course: studentObj.course, 
        pcIndex: pcIndex 
    };
    return fetch(GAS_URL, {
        method: 'POST', mode: 'no-cors',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
    });
}

/**
 * 3. 受講生登録（スプレッドシートの「受講生」シートへ保存）
 */
async function registerStudent() {
    const name = document.getElementById('s_name').value;
    const id = document.getElementById('s_id').value;
    const birth = document.getElementById('s_birth').value;
    const limit = document.getElementById('s_limit').value;
    const course = document.getElementById('s_course').value;
    if(!name || !id || !birth) return alert("未入力があります");

    // 受講生登録用のデータ（type: "student" を付ける）
    const studentData = { 
        type: "student", 
        id, name, 
        birthday: birth, 
        limit: Number(limit), 
        course,
        fixedDay: document.getElementById('s_fixed_day').value, // 追加
        fixedTime: document.getElementById('s_fixed_time').value // 追加
    };

    const overlay = document.getElementById('loadingOverlay');
    if(overlay) overlay.style.display = 'flex';

    try {
        await fetch(GAS_URL, { 
            method: 'POST', mode: 'no-cors', 
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(studentData) 
        });
        alert("受講生情報を保存しました。");
        await syncData(); // 全データを再取得して画面更新
    } catch(e) {
        alert("登録に失敗しました。");
    } finally {
        if(overlay) overlay.style.display = 'none';
    }
}

/**
 * 4. 受講生：予約実行
 */
async function makeBooking() {
    const selected = Array.from(document.querySelectorAll('input[name="slotTime"]:checked')).map(cb => cb.value);
    if(!selected.length || !selectedDateStr) return alert("時間を選択してください。");

    // 【月間上限チェック】
    const targetMonth = selectedDateStr.substring(0, 7); // "YYYY-MM" 形式
    // 現在の月の予約数をカウント
    const currentMonthResCount = reservations.filter(r => 
        r.studentId === currentStudent.id && 
        r.date.startsWith(targetMonth)
    ).length;

    if(!confirm("予約を確定しますか？")) return;

    document.getElementById('loadingOverlay').style.display = 'flex';
    try {
        await syncData(); // 予約直前に最新の予約リスト（固定枠含む）を取得
        const month = selectedDateStr.substring(0, 7);
        const count = reservations.filter(r => String(r.studentId) === String(currentStudent.id) && r.date.startsWith(month)).length;
        if (count + selected.length > Number(currentStudent.limit)) {
            throw new Error(`受講上限（${currentStudent.limit}回）を超えます。\n現在:${count}回、追加:${selected.length}回`);
        }
        for(let t of selected) {
            const timeNorm = normalizeTime(t);
            const resInSlot = reservations.filter(r => r.date === selectedDateStr && r.time === timeNorm);
            let assignedPc = -1;
            for(let p=0; p < STUDENT_CAPACITY; p++) {
                if(!resInSlot.some(r => Number(r.pcIndex) === p)) { assignedPc = p; break; }
            }
            if(assignedPc === -1) throw new Error(`${t} は満席です`);
            await sendToGoogleSheet(selectedDateStr, timeNorm, "予約", currentStudent, assignedPc);
        }
        await syncData();
        alert("予約完了しました");
        document.getElementById('timeSelectionArea').style.display = 'none';
        updateUI();
    } catch (e) { alert(e.message || "送信エラー"); }
    finally { document.getElementById('loadingOverlay').style.display = 'none'; }
}

/**
 * 5. 週間予約状況（座席表）の描画
 */
function renderWeeklySchedule() {
    const container = document.getElementById('weeklyScheduleContainer');
    const rangeDisplay = document.getElementById('adminWeekRangeDisplay');
    if(!container) return;
    container.innerHTML = '';
    const resMap = new Map();
    reservations.forEach(r => {
        if(!resMap.has(r.date)) resMap.set(r.date, []);
        resMap.get(r.date).push(r);
    });

    const tempDate = new Date(adminViewDate);
    const day = tempDate.getDay(); 
    const diffToTue = (day === 0) ? -5 : (day === 1) ? -6 : (2 - day);
    const tuesdayDate = new Date(tempDate);
    tuesdayDate.setDate(tempDate.getDate() + diffToTue);
    if(rangeDisplay) rangeDisplay.innerText = `${normalizeDate(tuesdayDate)} 〜 ${normalizeDate(new Date(tuesdayDate.getTime() + 4*24*60*60*1000))}`;

    let totalHtml = '';

    for(let i=0; i<5; i++) {
        const targetDate = new Date(tuesdayDate);
        targetDate.setDate(tuesdayDate.getDate() + i);
        const dateStr = normalizeDate(targetDate);
        const dayLabel = ['日','月','火','水','木','金','土'][targetDate.getDay()];
        
        let html = `<div class="day-schedule-wrapper"><h3 class="day-title">${dateStr} (${dayLabel})</h3>`;
        html += `<table class="schedule-table"><thead><tr><th>時間</th>`;
        for(let pc=1; pc<=ADMIN_CAPACITY; pc++) { html += `<th>PC${pc}</th>`; }
        html += `</tr></thead><tbody>`;


        TIMES.forEach(timeLabel => {
            const tNorm = normalizeTime(timeLabel);
            const rowClass = (tNorm === "09:00" || tNorm === "12:00") ? "row-gray" : "";
            html += `<tr class="${rowClass}"><td>${timeLabel}</td>`;
            
            //const resInSlot = reservations.filter(r => r.date === dateStr && r.time === tNorm);
            // 全体（reservations）ではなく、その日のリスト（resMap）から取得します
            const dayRes = resMap.get(dateStr) || [];
            const resInSlot = dayRes.filter(r => r.time === tNorm);
    
            const manToManRes = resInSlot.find(r => String(r.course).includes("マンツーマンコース"));

            if (manToManRes) {
                // マンツーマン予約がある場合、その行の全セルを cell-pc-man クラスにする
                for(let pc=0; pc<ADMIN_CAPACITY; pc++) {
                    // 名前は予約されたpcIndexと一致する列にだけ表示する
                    const displayName = (pc === Number(manToManRes.pcIndex)) ? manToManRes.name : "";
                    html += `<td class="cell-pc-man" onclick="handleAdminCellClick('${dateStr}', '${timeLabel}', ${pc}, '${manToManRes.studentId}')">${displayName}</td>`;
                }
            } else {
                // 通常の描画
                let rowSeats = new Array(ADMIN_CAPACITY).fill(null);
                resInSlot.forEach(r => { if (r.pcIndex !== null && r.pcIndex >= 0 && r.pcIndex < ADMIN_CAPACITY) rowSeats[r.pcIndex] = r; });
                resInSlot.forEach(r => { if (!rowSeats.includes(r)) { const emptyIdx = rowSeats.indexOf(null); if (emptyIdx !== -1) rowSeats[emptyIdx] = r; } });
                
                for(let pc=0; pc<ADMIN_CAPACITY; pc++) {
                    const r = rowSeats[pc];
                    if(r) {
                        // let cls = String(r.course).includes("パソコン") ? "cell-pc-blue" : (String(r.course).includes("プログラミング") ? "cell-pc-red" : "cell-pc-default");

                        let cls = "";
                        if (String(r.course).includes("休み")) {
                            cls = "cell-holiday"; // 休み専用の色
                        } else if (String(r.course).startsWith("検定：")) {
                            cls = "cell-exam";
                        } else if (String(r.course).includes("マンツーマンコース")) {
                            cls = "cell-pc-man";
                        } else {
                            // 既存の判定
                            cls = String(r.course).includes("パソコン") ? "cell-pc-blue" : (String(r.course).includes("プログラミング") ? "cell-pc-red" : "cell-pc-default");
                        }

                        html += `<td class="${cls}" onclick="handleAdminCellClick('${dateStr}', '${timeLabel}', ${pc}, '${r.studentId}')">${r.name}</td>`;
                    } else {
                        html += `<td onclick="handleAdminCellClick('${dateStr}', '${timeLabel}', ${pc}, null)"></td>`;
                    }
                }
            }
            html += `</tr>`;
        });
        //container.innerHTML += html + `</tbody></table></div>`;
        totalHtml += html + "</tbody></table></div>";
    }
    container.innerHTML = totalHtml;
}

/**
 * 以降、UI制御・補助関数
 */
async function handleAdminCellClick(date, time, pcIndex, studentId) {
    if(studentId) {
        const tNorm = normalizeTime(time);
        //const r = reservations.find(res => res.date === date && res.time === tNorm && res.studentId === studentId);
        // const r = reservations.find(res => 
        //     res.date === date && 
        //     res.time === tNorm && 
        //     String(res.studentId).trim() === String(studentId).trim() && 
        //     Number(res.pcIndex) === Number(pcIndex)
        // );
        const r = reservations.find(res => {
            // 日付、時間、受講生IDが一致していることを大前提とする
            const baseMatch = res.date === date && 
                            res.time === tNorm && 
                            String(res.studentId).trim() === String(studentId).trim();

            if (!baseMatch) return false;

            // 検定試験（EXAM）や休み等の場合は、複数の同一IDが同じ枠に存在するためPC番号までチェックする
            if (String(studentId).trim() === "EXAM") {
                return Number(res.pcIndex) === Number(pcIndex);
            }

            // 通常の受講生の場合は、その時間にそのIDの人は1人しかいないため、
            // 表示上のPC番号（pcIndex）がズレていても、本人であれば特定（find）に成功させます
            return true;
        });
        if(!r) return;
        if(confirm(`${r.name} 様の予約を取り消しますか？`)) {
            document.getElementById('loadingOverlay').style.display = 'flex';
            await sendToGoogleSheet(r.date, r.time, "取消", {id: r.studentId, name: r.name, course: r.course}, r.pcIndex);
            await syncData();
            renderWeeklySchedule();
            document.getElementById('loadingOverlay').style.display = 'none';
        }
    } else {
        selectedCell = { date, time, pcIndex };
        document.getElementById('modalDate').innerText = date;
        document.getElementById('modalTime').innerText = time;
        document.getElementById('modalPc').innerText = pcIndex + 1;

        // 受講生リストをセット
        document.getElementById('modalStudentSelect').innerHTML = '<option value="">受講生を選択</option>' + students.map(s => `<option value="${s.id}">${s.name}</option>`).join('');
        // 検定リストをセット
        // document.getElementById('modalExamSelect').innerHTML = '<option value="">-- 検定を選択 --</option>' + examNames.map(n => `<option value="${n}">${n}</option>`).join('');
        const modalExamSelect = document.getElementById('modalExamSelect');
        if (modalExamSelect) {
            modalExamSelect.innerHTML = '<option value="">-- 検定を選択 --</option>' + 
                examNames.map(n => `<option value="${n}">${n}</option>`).join('');
        }
        const studentArea = document.getElementById('studentSelectionArea');
        const examArea = document.getElementById('examSelectionArea');

        // --- PCごとの表示切り替えロジック ---
        
        if (pcIndex >= 5) { 
            // PC6(index 5)・PC7(index 6) の場合：検定のみ表示
            studentArea.style.display = 'none';
            examArea.style.display = 'block';
            document.getElementById('modalStudentSelect').value = "";
        } else if (pcIndex >= 2) { 
            // PC3(index 2) 〜 PC5(index 4) の場合：両方表示（受講生 or 検定）
            studentArea.style.display = 'block';
            examArea.style.display = 'block';
        } else {
            // PC1(index 0)・PC2(index 1) の場合：受講生のみ表示
            studentArea.style.display = 'block';
            examArea.style.display = 'none';
            document.getElementById('modalExamSelect').value = "";
        }

        document.getElementById('adminBookingModal').style.display = 'block';
    }
}

async function executeAdminBooking() {
    const sid = document.getElementById('modalStudentSelect').value;
    // const examName = document.getElementById('modalExamSelect') ? document.getElementById('modalExamSelect').value : "";
    const examName = document.getElementById('modalExamSelect').value;
    if(!sid && !examName) return alert("受講生または検定を選択してください");

    const date = selectedCell.date;
    const time = normalizeTime(selectedCell.time);
    const pcIdx = selectedCell.pcIndex;
    const resInSlot = reservations.filter(r => r.date === date && r.time === time);

    const isHolidayInSlot = resInSlot.some(r => String(r.course).includes("休み"));
    if (isHolidayInSlot) {
        return alert("この時間は「休み」に設定されているため、追加の予約はできません。");
    }

    // 制約チェック
    const hasManToMan = resInSlot.some(r => String(r.course).includes("マンツーマンコース"));
    const hasProgramming = resInSlot.some(r => String(r.course).includes("プログラミング"));
    const hasExam = resInSlot.some(r => String(r.course).startsWith("検定："));

    if (examName) {
        // --- 検定試験の優先順位チェックロジック ---
        // 優先順位: PC6(5), PC7(6), PC5(4), PC4(3), PC3(2)
        const priorityOrder = [5, 6, 4, 3, 2];
        const isOccupied = (idx) => resInSlot.some(r => Number(r.pcIndex) === idx);

        // クリックされたPCより優先度が高いPCが空いていないか確認
        for (let p of priorityOrder) {
            if (p === pcIdx) break; // 自分の番まで来たらOK
            if (!isOccupied(p)) {
                return alert(`PC${p + 1} が空いています。PC${p + 1} から順に使用してください。`);
            }
        }

        // コース制約チェック
        const hasManToMan = resInSlot.some(r => String(r.course).includes("マンツーマンコース"));
        const hasProgramming = resInSlot.some(r => String(r.course).includes("プログラミング"));
        if (hasManToMan || hasProgramming) return alert("マンツーマンまたはプログラミングの予約があるため、検定は登録できません。");
    } else {
        const st = students.find(s => s.id === sid);
        if (hasExam && (st.course.includes("マンツーマンコース") || st.course.includes("プログラミング"))) {
            return alert("検定試験が登録されているため、マンツーマンまたはプログラミングは予約できません。");
        }
    }

    const overlay = document.getElementById('loadingOverlay');
    if(overlay) overlay.style.display = 'flex';

    try {
        let studentObj, courseName;
        if (examName) {
            studentObj = { id: "EXAM", name: examName };
            courseName = "検定：" + examName;
        } else {
            studentObj = students.find(s => s.id === sid);
            courseName = studentObj.course;
        }
        
        await sendToGoogleSheet(date, time, "予約", { ...studentObj, course: courseName }, pcIdx);
        // await sendToGoogleSheet(date, time, "予約", { ...studentObj, course: courseName }, selectedCell.pcIndex);
        await syncData(); closeModal(); renderWeeklySchedule();
    } catch(e) { alert("送信エラー"); }
    finally { if(overlay) overlay.style.display = 'none'; }
}

// async function executeAdminBooking() {
//     const sid = document.getElementById('modalStudentSelect').value;
//     if(!sid) return;
//     const st = students.find(s => s.id === sid);
//     document.getElementById('loadingOverlay').style.display = 'flex';
//     try {
//         await sendToGoogleSheet(selectedCell.date, selectedCell.time, "予約", st, selectedCell.pcIndex);
//         await syncData(); closeModal(); renderWeeklySchedule();
//     } catch(e) { alert("送信エラー"); }
//     finally { document.getElementById('loadingOverlay').style.display = 'none'; }
// }

function handleLogin() {
    const id = document.getElementById('loginId').value;
    const birth = document.getElementById('loginBirth').value;
    const user = students.find(s => s.id === id && s.birthday === birth);
    if (user) {
        currentStudent = user;
        document.getElementById('login-form').style.display = 'none';
        document.getElementById('student-actions').style.display = 'block';
        document.getElementById('displayUserName').innerText = user.name;
        updateUI();
    } else { alert("ログイン失敗"); }
}

function handleLogout() { currentStudent = null; document.getElementById('login-form').style.display = 'block'; document.getElementById('student-actions').style.display = 'none'; }
function updateUI() { if (!currentStudent) return; renderCalendar(); const m = normalizeDate(viewDate).substring(0, 7); renderMyReservations(m); }
function selectDate(dateStr) { selectedDateStr = dateStr; document.getElementById('selectedDateText').innerText = dateStr; document.getElementById('timeSelectionArea').style.display = 'block'; renderCalendar(); renderSlots(dateStr); }
function saveData() { localStorage.setItem('sch_v5_students', JSON.stringify(students)); localStorage.setItem('sch_v5_res', JSON.stringify(reservations)); }
function closeModal() { document.getElementById('adminBookingModal').style.display = 'none'; }
function changeAdminWeek(days) { adminViewDate.setDate(adminViewDate.getDate() + days); renderWeeklySchedule(); }
function changeMonth(diff) { viewDate.setMonth(viewDate.getMonth() + diff); renderCalendar(); const m = normalizeDate(viewDate).substring(0, 7); renderMyReservations(m); }

function switchAdminSubTab(subTab) {
    const views = ['admin-week-view', 'admin-month-view', 'admin-student-view'];
    views.forEach(v => { const el = document.getElementById(v); if(el) el.style.display = (v === 'admin-' + subTab + '-view') ? 'block' : 'none'; });
    document.querySelectorAll('.admin-tabs button').forEach(b => b.classList.toggle('active', b.id === 'btn-tab-' + subTab));
    if(subTab === 'month') renderMonthlyList();
}

function renderMonthlyList() {
    const container = document.getElementById('monthlyListContainer'); if(!container) return;
    const year = adminListMonth.getFullYear(); const month = adminListMonth.getMonth() + 1;
    document.getElementById('adminListMonthDisplay').innerText = `${year}年${month}月`;
    const prefix = `${year}-${String(month).padStart(2,'0')}`;
    let html = `<table class="monthly-list-table"><thead><tr><th>受講生名</th><th>枠</th><th>残</th>`;
    for(let i=1; i<=20; i++) { html += `<th>${i}</th>`; }
    html += `</tr></thead><tbody>`;
    students.forEach(s => {
        const myMonthRes = reservations.filter(r => 
            String(r.studentId).trim() === String(s.id).trim() && 
            r.date.startsWith(prefix)
        ).sort((a,b) => a.date.localeCompare(b.date)|| a.time.localeCompare(b.time));
        html += `<tr><td style="text-align:left;">${s.name}</td><td>${s.limit}</td><td>${s.limit - myMonthRes.length}</td>`;
        for(let i=0; i<20; i++) { 
            if(myMonthRes[i]) { 
                const dP = myMonthRes[i].date.split('-'); 
                html += `<td>${Number(dP[1])}/${Number(dP[2])}</td>`; 
            } else { 
                html += `<td></td>`; 
            } 
        }
        html += `</tr>`;
    });
    container.innerHTML = html + `</tbody></table>`;
}

function changeAdminListMonth(diff) { adminListMonth.setMonth(adminListMonth.getMonth() + diff); renderMonthlyList(); }

function renderStudentList() {
    const list = document.getElementById('studentList'); if(!list) return;
    list.innerHTML = "<h4>登録受講生</h4>" + students.map(s => `
        <div class="list-item"><span><b>${s.name}</b> (${s.id})</span><button class="btn-outline btn-sm btn-danger" onclick="deleteStudent('${s.id}')">削除</button></div>
    `).join('');
}

function deleteStudent(id) { if(confirm("削除しますか？")){ students = students.filter(s => s.id !== id); saveData(); renderStudentList(); } }

async function cancelBooking(date, time) {
    if(!confirm("この予約を取り消しますか？")) return;
    const overlay = document.getElementById('loadingOverlay');
    if(overlay) overlay.style.display = 'flex';
    try {
        const target = reservations.find(res => res.studentId === currentStudent.id && res.date === date && res.time === time);
        await sendToGoogleSheet(date, time, "取消", currentStudent, target ? target.pcIndex : null);
        await syncData(); updateUI();
    } catch(e) { alert("取消エラー"); }
    finally { if(overlay) overlay.style.display = 'none'; }
}

function renderMyReservations(monthPrefix) {
    const list = document.getElementById('myReservations'); if(!list) return;
    const todayStr = normalizeDate(new Date());
    const myRes = reservations.filter(r => r.studentId === currentStudent.id && r.date.startsWith(monthPrefix)).sort((a,b) => a.date.localeCompare(b.date) || a.time.localeCompare(b.time));
    list.innerHTML = myRes.map(r => `
        <div class="list-item"><div><strong>${r.date.substring(5)} ${r.time}</strong> <small>${r.course}</small></div>
        ${r.date >= todayStr ? `<button class="btn-outline btn-sm btn-danger" onclick="cancelBooking('${r.date}', '${r.time}')">取消</button>` : '<small>終了</small>'}</div>
    `).join('') || '<p style="text-align:center; color:#999;">この月の予約はありません</p>';
}

function renderSlots(date) {
    const container = document.getElementById('slotContainer'); if(!container) return; container.innerHTML = '';
    const dayOfWeek = new Date(date.replace(/-/g, '/')).getDay();
    const daysRes = reservations.filter(r => r.date === date);

    TIMES.forEach(t => {
        const tNorm = normalizeTime(t);
        const resInSlot = daysRes.filter(r => r.time === tNorm);
        const count = resInSlot.length;
        const remain = STUDENT_CAPACITY - count;
        const booked = resInSlot.some(r => r.studentId === currentStudent.id);
        
        // 1. 特殊な予約（マンツーマン・検定）の有無をチェック
        const hasManToMan = resInSlot.some(r => String(r.course).includes("マンツーマンコース"));
        const hasExamInSlot = resInSlot.some(r => String(r.course).startsWith("検定："));
        

        // 2. 基本的なブロック条件（時間外、自身が予約済、満席など）
        let isSpecialBlock = (tNorm === "09:00" || tNorm === "12:00");
        let isBlocked = isSpecialBlock || (tNorm === "18:00" && (dayOfWeek === 2 || dayOfWeek === 6)) || (tNorm === "19:00" && dayOfWeek !== 3) || booked || remain <= 0;

        const isHoliday = resInSlot.some(r => String(r.course).includes("休み"));

        // 基本的なブロック条件に「休み」を追加
        if (isHoliday) {
            isBlocked = true;
        }

        // 3. コースごとの制約ロジックを追加
        const myCourse = String(currentStudent.course);

        if (hasManToMan) {
            // マンツーマンが入っている時間は、誰であっても予約不可
            isBlocked = true;
        }

        if (hasExamInSlot) {
            // 検定が入っている時間は、プログラミング教室とマンツーマンコースの人は予約不可
            if (myCourse.includes("プログラミング") || myCourse.includes("マンツーマンコース")) {
                isBlocked = true;
            }
        }

        // 警告表示（残りわずか）の判定：ブロックされていない場合のみ
        let isWarn = !isBlocked && !booked && remain > 0 && remain <= 2;
        
        const div = document.createElement('div');
        div.className = `slot-card ${isBlocked && !isWarn ? 'full' : (isWarn ? 'warn-block' : '')}`;
        
        // アイコン表示の決定
        let txt = isSpecialBlock ? "－" : (booked ? "済" : (isWarn ? "▲" : (isBlocked ? "×" : "〇")));
        
        div.innerHTML = `<strong>${t}</strong><br><span class="badge ${isWarn?'bg-warn':(isBlocked?'bg-ng':'bg-ok')}">${txt}</span>${!isBlocked && !isWarn ? `<input type="checkbox" name="slotTime" value="${t}">` : ''}`;
        
        if(isWarn) div.onclick = () => alert("残りわずかのため、教室へお問い合わせください。");
        else if(!isBlocked) div.onclick = (e) => { if(e.target.type !== 'checkbox') { const cb = div.querySelector('input'); cb.checked = !cb.checked; } };
        container.appendChild(div);
    });
}
function renderCalendar() {
    const grid = document.getElementById('calendarGrid'); if(!grid) return; grid.innerHTML = '';
    const year = viewDate.getFullYear(); const month = viewDate.getMonth();
    const disp = document.getElementById('calendarMonthDisplay'); if(disp) disp.innerText = `${year}年${month + 1}月`;
    ['日','月','火','水','木','金','土'].forEach(d => { const el = document.createElement('div'); el.className = 'cal-day-label'; el.innerText = d; grid.appendChild(el); });
    const firstDay = new Date(year, month, 1).getDay(); const lastDate = new Date(year, month+1, 0).getDate();
    for(let i=0; i<firstDay; i++) { grid.appendChild(document.createElement('div')); }
    for(let date=1; date<=lastDate; date++) {
        const dStr = normalizeDate(new Date(year, month, date));
        const el = document.createElement('div'); el.className = 'cal-date'; el.innerText = date;
        if(reservations.some(r => r.studentId === currentStudent.id && r.date === dStr)) el.classList.add('has-res');
        if(dStr === normalizeDate(new Date())) el.classList.add('today');
        if(dStr === selectedDateStr) el.classList.add('selected');
        if(dStr < normalizeDate(new Date())) el.classList.add('disabled'); else el.onclick = () => selectDate(dStr);
        grid.appendChild(el);
    }
}

// 検定登録
async function registerExamName() {
    const name = document.getElementById('newExamName').value;
    if(!name) return;
    const overlay = document.getElementById('loadingOverlay');
    if(overlay) overlay.style.display = 'flex';
    try {
        await fetch(GAS_URL, {
            method: 'POST', mode: 'no-cors',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ type: "examMaster", name: name })
        });
        document.getElementById('newExamName').value = '';
        await syncData();
    } finally {
        if(overlay) overlay.style.display = 'none';
    }
}

// 受講生選択ドロップダウンの中身を更新する
function updateStudentEditDropdown() {
    const sel = document.getElementById('s_select_list');
    if (!sel) return;
    const currentVal = sel.value; // 現在の選択を維持
    sel.innerHTML = '<option value="">-- 新規登録（または直接入力） --</option>' +
        students.map(s => `<option value="${s.id}">${s.name} (${s.id})</option>`).join('');
    sel.value = currentVal;
}

// ドロップダウンで受講生を選んだら入力欄に値をセットする
function handleStudentEditSelect() {
    const id = document.getElementById('s_select_list').value;
    if (!id) {
        // 新規登録が選ばれたら入力を空にする
        document.getElementById('s_name').value = '';
        document.getElementById('s_id').value = '';
        document.getElementById('s_birth').value = '';
        document.getElementById('s_limit').value = '4';
        document.getElementById('s_course').value = 'パソコン教室';
        document.getElementById('s_fixed_day').value = '';
        document.getElementById('s_fixed_time').value = '';
        return;
    }
    
    // 選択された受講生の情報を探して各入力欄に入れる
    const s = students.find(x => String(x.id) === String(id));
    if (s) {
        document.getElementById('s_name').value = s.name;
        document.getElementById('s_id').value = s.id;
        document.getElementById('s_birth').value = s.birthday;
        document.getElementById('s_limit').value = s.limit;
        document.getElementById('s_course').value = s.course;
        document.getElementById('s_fixed_day').value = s.fixedDay || "";
        document.getElementById('s_fixed_time').value = s.fixedTime || "";
    }
}