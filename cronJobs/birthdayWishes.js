const cron = require('node-cron');
const sendEmail = require('./../utils/email');
const xlsx = require('xlsx') 

cron.schedule('0 10 * * *', async () => {
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

        // Find users having birthday today
        const birthdayPeople = data.filter(user => user.dob.slice(5) === now);
    
        // Send birthday wish to the users
        for(const user of birthdayPeople) {
            const userFname = user.name.split(' ')[0];
            await sendEmail({
                email: user.email,
                subject: `Happy Birthday, ${userFname}!`,
                message: `Dear ${user.name},

Wishing you a very happy birthday! 🎉 May your day be filled with joy, laughter, and celebration. We are grateful to have you as part of our team and hope this year brings you continued success and happiness.

Enjoy your special day!`
            });
            console.log(`Birthday wish sent to ${user.name}`);
        }

    } catch (error) {
        console.log(error)
    }
})