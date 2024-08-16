const cron = require('node-cron');
const sendEmail = require('./../utils/email');
const xlsx = require('xlsx')

cron.schedule('47 11 * * *', async () => {
    try {
        const now = new Date().toJSON().slice(5, 10);
        const today = new Date().toJSON().slice(0, 10);

        // Reading our excel file 
        const employees = xlsx.readFile('./employee_details.xlsx')
        const festivals = xlsx.readFile('./festivals.xlsx')

        let data = []
        let data2 = []

        const employee_sheets = employees.SheetNames;
        const festival_sheets = festivals.SheetNames;

        for (let i = 0; i < employee_sheets.length; i++) {
            const temp = xlsx.utils.sheet_to_json(employees.Sheets[employees.SheetNames[i]])
            temp.forEach((res) => {
                data.push(res)
            })
        }

        for (let i = 0; i < festival_sheets.length; i++) {
            const temp = xlsx.utils.sheet_to_json(festivals.Sheets[festivals.SheetNames[i]])
            temp.forEach((res) => {
                data2.push(res)
            })
        }

        // Birthday 
        const birthdayPeople = data.filter(user => user.dob.slice(5) === now);

        for (const user of birthdayPeople) {
            const userFname = user.name.split(' ')[0];
            await sendEmail({
                email: user.email,
                subject: `Happy Birthday, ${userFname}!`,
                message: `Dear ${user.name},

Wishing you a very happy birthday! 🎉 May your day be filled with joy, laughter, and celebration. We are grateful to have you as part of our team and hope this year brings you continued success and happiness.

Enjoy your special day!`
            });
            // console.log(`Birthday wish sent to ${user.name}`);
        }

        // Work anniversary 
        const workAnniversary = data.filter(user => user.joining_date.slice(5) === now);

        for (const user of workAnniversary) {
            const userFname = user.name.split(' ')[0];
            const years = new Date().getFullYear() - new Date(user.joining_date).getFullYear();
            let abb = '';
            if (years === 1) {
                abb = 'st'
            }
            else if (years === 2) {
                abb = 'nd'
            }
            else if (years === 3) {
                abb = 'rd'
            }
            else {
                abb = 'th'
            }
            await sendEmail({
                email: user.email,
                subject: `Happy Work Anniversary, ${userFname}!`,
                message: `Dear ${user.name},

Congratulations on your ${years}${abb} work anniversary! 🎉

Your dedication, hard work, and contributions have been instrumental to our success. We are grateful to have you as part of our team and look forward to many more successful years together.

Thank you for your continued commitment, and here’s to celebrating more milestones in the future!`
            });
            // console.log(`Work anniversary wish sent to ${user.name}`);
        }

        // Festival wishes
        const date_format_festivals = data2.map(festival => { return { festival, date: excelDateToJSDate(festival.date) } })
        const today_festivals = date_format_festivals.filter(festival => festival.date === today);

        for (const festival of today_festivals) {
            // console.log(festival.festival.festival_name)
            for (const user of data) {
                const userFname = user.name.split(' ')[0];
                await sendEmail({
                    email: user.email,
                    subject: `Happy ${festival.festival.festival_name}, ${userFname}!`,
                    message: `Dear ${user.name},

As ${festival.festival.festival_name} approaches, I wanted to extend my heartfelt wishes to you and your loved ones. May this festive season bring you joy, peace, and prosperity.

Let’s take this opportunity to celebrate, reflect, and recharge. I hope you enjoy the festivities and create beautiful memories with those who matter most.

Wishing you a wonderful ${festival.festival.festival_name}!`
                });
                // console.log(`Festival wish sent to ${user.name}`);
            }
        }
    } catch (error) {
        console.log(error)
    }
})

function excelDateToJSDate(serial) {
    // Excel's date system is based on the 1900 date system in Windows
    // Excel's date 1 corresponds to 1900-01-01 (for Windows)
    var excelEpoch = new Date(1900, 0, 1);
    // Adjust for the Excel leap year bug, Excel considers 1900 as a leap year (it is not)
    var dateOffset = serial > 59 ? serial - 1 : serial;
    var days = dateOffset; // Adjust for the offset between JavaScript and Excel days
    // +1 date
    return new Date(excelEpoch.getTime() + days * 24 * 60 * 60 * 1000).toJSON().slice(0, 10);
}