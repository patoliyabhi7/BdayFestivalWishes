const cron = require('node-cron');
const sendEmail = require('./../utils/email');
const xlsx = require('xlsx')

cron.schedule('0 10 * * *', async () => {
    try {
        const now = new Date().toJSON().slice(5, 10);
        const today = new Date().toJSON().slice(0, 10);

        // Reading our excel file 
        const employees = xlsx.readFile('./employee.xlsx')
        const festivals = xlsx.readFile('./festivals.xlsx')

        let data = []
        let data2 = []

        const employee_sheets = employees.SheetNames;
        const festival_sheets = festivals.SheetNames;

        for (let i = 0; i < employee_sheets.length; i++) {
            const temp = xlsx.utils.sheet_to_json(employees.Sheets[employees.SheetNames[i]])
            data = data.concat(temp);
        }

        for (let i = 0; i < festival_sheets.length; i++) {
            const temp = xlsx.utils.sheet_to_json(festivals.Sheets[festivals.SheetNames[i]])
            data2 = data2.concat(temp);
        }

        // Birthday 
        const date_format_bday = data.map(user => {
            return { user, date: excelDateToJSDate(user.dob) };
        });
        const birthdayPeople = date_format_bday.filter(user => user.date.slice(5) === now);

        for (const entry of birthdayPeople) {
            const { user } = entry;
            if (user.email) {
                const userFname = user.name.split(' ')[0];
                await sendEmail({
                    email: user.email,
                    subject: `Happy Birthday, ${userFname}!`,
                    message: `Dear ${user.name},

Wishing you a very happy birthday! 🎉 May your day be filled with joy, laughter, and celebration. We are grateful to have you as part of our team and hope this year brings you continued success and happiness.

Enjoy your special day!`
                });
            } else {
                console.log(`No email found for user: ${user.name}`);
            }
        }

        // Work anniversary 
        const date_format_anniversary = data.map(user => {
            return { user, date: excelDateToJSDate(user.joining_date) };
        });
        const workAnniversary = date_format_anniversary.filter(user => user.date.slice(5) === now);

        for (const entry of workAnniversary) {
            console.log(entry)
            const { user } = entry;
            if (user.email) {
                const userFname = user.name.split(' ')[0];
                const years = new Date().getFullYear() - new Date(entry.date).getFullYear();
                const abb = years === 1 ? 'st' : years === 2 ? 'nd' : years === 3 ? 'rd' : 'th';
                await sendEmail({
                    email: user.email,
                    subject: `Happy Work Anniversary, ${userFname}!`,
                    message: `Dear ${user.name},

Congratulations on your ${years}${abb} work anniversary! 🎉

Your dedication, hard work, and contributions have been instrumental to our success. We are grateful to have you as part of our team and look forward to many more successful years together.

Thank you for your continued commitment, and here’s to celebrating more milestones in the future!`
                });
            } else {
                console.log(`No email found for user: ${user.name}`);
            }
        }

        // Festival wishes
        const date_format_festivals = data2.map(festival => {
            return { festival, date: excelDateToJSDate(festival.date) };
        });
        const today_festivals = date_format_festivals.filter(festival => festival.date === today);

        for (const festival of today_festivals) {
            for (const user of data) {
                if (user.email) {
                    const userFname = user.name.split(' ')[0];
                    await sendEmail({
                        email: user.email,
                        subject: `Happy ${festival.festival.festival_name}, ${userFname}!`,
                        message: `Dear ${user.name},

As ${festival.festival.festival_name} approaches, I wanted to extend my heartfelt wishes to you and your loved ones. May this festive season bring you joy, peace, and prosperity.

Let’s take this opportunity to celebrate, reflect, and recharge. I hope you enjoy the festivities and create beautiful memories with those who matter most.

Wishing you a wonderful ${festival.festival.festival_name}!`
                    });
                } else {
                    console.log(`No email found for user: ${user.name}`);
                }
            }
        }
    } catch (error) {
        console.log("Error in cron job:", error);
    }
});

function excelDateToJSDate(serial) {
    var excelEpoch = new Date(1900, 0, 1);
    var dateOffset = serial > 59 ? serial - 1 : serial;
    var days = dateOffset;
    return new Date(excelEpoch.getTime() + days * 24 * 60 * 60 * 1000).toJSON().slice(0, 10);
}
