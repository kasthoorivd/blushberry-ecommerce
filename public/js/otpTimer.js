document.addEventListener("DOMContentLoaded", function () {
  startOtpTimer("timer", "resendBtn", 60);
});

function startOtpTimer(timerId, buttonId, duration = 60) {

  const timer = document.getElementById(timerId);
  const resendBtn = document.getElementById(buttonId);

  if (!timer || !resendBtn) {
    console.log("Timer or Button not found");
    return;
  }

  let endTime = localStorage.getItem("otpEndTime");

  if (!endTime) {
    endTime = Date.now() + duration * 1000;
    localStorage.setItem("otpEndTime", endTime);
  } else {
    endTime = parseInt(endTime);
  }

  resendBtn.disabled = true;
  resendBtn.classList.remove("enabled");

  const interval = setInterval(() => {

    const timeLeft = Math.floor((endTime - Date.now()) / 1000);

    if (timeLeft <= 0) {
      clearInterval(interval);
      timer.innerText = "";
      resendBtn.disabled = false;
      resendBtn.classList.add("enabled");
      localStorage.removeItem("otpEndTime");
      return;
    }

    timer.innerText = `Resend OTP in ${timeLeft}s`;

  }, 1000);
}