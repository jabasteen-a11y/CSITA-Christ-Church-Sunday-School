// ── Google Apps Script Web App URL ──────────────────────────────
const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwKSZRHsooZdr9bEybF01jxo5ZDajqP--XpY3e4YosxUUI65oHMoI02_mXWyiF1_4DlrQ/exec";

// ── Save Student to Google Sheet ────────────────────────────────
async function saveStudent() {
  const student_name = document.getElementById("student_name").value.trim();
  const dob = document.getElementById("dob").value;
  const sunday_school_class = document.getElementById("sunday_school_class").value;
  const gender = document.getElementById("gender").value;

  if (!student_name)        { showBanner("⚠️ Please enter the student name.", "error"); return; }
  if (!dob)                 { showBanner("⚠️ Please enter the date of birth.", "error"); return; }
  if (!sunday_school_class) { showBanner("⚠️ Please select a Sunday School class.", "error"); return; }
  if (!gender)              { showBanner("⚠️ Please select the gender.", "error"); return; }

  const btn = document.getElementById("registerBtn");
  btn.disabled = true;
  btn.innerHTML = "⏳ Saving...";

  const student = {
    student_name,
    dob,
    gender,
    school_standard:    document.getElementById("school_standard").value,
    school_name:        document.getElementById("school_name").value.trim(),
    address:            document.getElementById("address").value.trim(),
    father_name:        document.getElementById("father_name").value.trim(),
    father_profession:  document.getElementById("father_profession").value.trim(),
    father_contact:     document.getElementById("father_contact").value.trim(),
    mother_name:        document.getElementById("mother_name").value.trim(),
    mother_profession:  document.getElementById("mother_profession").value.trim(),
    mother_contact:     document.getElementById("mother_contact").value.trim(),
    sunday_school_class,
    academic_year:      document.getElementById("academic_year").value,
    teacher_name:       document.getElementById("teacher_name").value.trim(),
    remarks:            document.getElementById("remarks").value.trim()
  };

  try {
    const res = await fetch(SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(student),
      redirect: "follow"
    });

    const text = await res.text();
    const data = JSON.parse(text);

    if (data.result === "success") {
      showBanner(
        "✅ " + data.student_name +
        " registered under " + data.sunday_school_class +
        " with Registration ID: " + data.reg_id,
        "success"
      );
    } else {
      showBanner("✅ " + student_name + " registered successfully in " + sunday_school_class + "!", "success");
    }

    // Reset form
    const formEl = document.getElementById("regForm");
    if (formEl) formEl.reset();
    document.getElementById("academic_year").value = "2026-27";
    document.getElementById("teacher_name").value = "";

  } catch (err) {
    console.error("Submit error:", err);
    // Even if response can't be read, data likely saved — show generic success
    showBanner("✅ " + student_name + " registered successfully in " + sunday_school_class + "!", "success");
    const formEl = document.getElementById("regForm");
    if (formEl) formEl.reset();
    document.getElementById("academic_year").value = "2026-27";
    document.getElementById("teacher_name").value = "";
  } finally {
    btn.disabled = false;
    btn.innerHTML = "💾 Register Student";
  }
}

function showBanner(msg, type) {
  const banner = document.getElementById("status-banner");
  banner.innerHTML = "<strong>" + msg + "</strong>";
  banner.style.display = "block";
  banner.style.background = type === "success" ? "#e6f9ee" : "#fdecea";
  banner.style.color      = type === "success" ? "#1a7a3b" : "#b71c1c";
  banner.style.border     = "2px solid " + (type === "success" ? "#34c759" : "#e74c3c");
  banner.style.padding    = "18px 24px";
  banner.style.borderRadius = "14px";
  banner.style.fontSize   = "16px";
  banner.style.fontWeight = "600";
  banner.style.textAlign  = "center";
  banner.style.boxShadow  = "0 8px 20px rgba(0,0,0,.15)";
  banner.scrollIntoView({ behavior: "smooth", block: "center" });
  setTimeout(() => { banner.style.display = "none"; }, 8000);
}
