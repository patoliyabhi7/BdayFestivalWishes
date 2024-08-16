const cron = require('node-cron');
const sendEmail = require('./../utils/email');
const xlsx = require('xlsx')

cron.schedule('30 16 10 * * *', async () => {
    try {
        const now = new Date().toJSON().slice(5, 10);

        // Reading our excel file 
        const file = xlsx.readFile('./employee_details.xlsx')

        let data = []

        const sheets = file.SheetNames;

        for (let i = 0; i < sheets.length; i++) {
            const temp = xlsx.utils.sheet_to_json(
                file.Sheets[file.SheetNames[i]])
            temp.forEach((res) => {
                data.push(res)
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

    } catch (error) {
        console.log(error)
    }
})