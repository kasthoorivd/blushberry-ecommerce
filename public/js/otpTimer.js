let otpInterval;

function startOtpTimer(timerId, buttonId, duration=60){
    const timer = document.getElementById(timerId);
    const resendBtn = document.getElementById(buttonId);
    if(!timer || !resendBtn) return;

    // clear previous interval
    if(otpInterval) clearInterval(otpInterval);

    // get endTime or set new
    let endTime = localStorage.getItem('otpEndTime');
    if(!endTime){
        endTime = Date.now() + duration*1000;
        localStorage.setItem('otpEndTime', endTime);
    } else {
        endTime = parseInt(endTime);
    }

    // enable/disable button based on cooldown
    if(Date.now() < endTime){
        resendBtn.disabled = true;
        resendBtn.classList.remove('enabled');
    } else {
        resendBtn.disabled = false;
        resendBtn.classList.add('enabled');
    }

    otpInterval = setInterval(() => {
        const timeLeft = Math.floor((endTime - Date.now())/1000);
        if(timeLeft <= 0){
            clearInterval(otpInterval);
            timer.innerText = "";
            resendBtn.disabled = false;
            resendBtn.classList.add('enabled');
            localStorage.removeItem('otpEndTime');
            return;
        }
        timer.innerText = `Resend OTP in ${timeLeft}s`;
    }, 1000);
}