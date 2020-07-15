require('chromedriver').path;
const {
    Builder,
    By,
    Key,
    util
} = require('selenium-webdriver');
const XLSX = require('xlsx');
const moment = require('moment');

readExcelData();

//read excel data impln
async function readExcelData() {

        const workbook = XLSX.readFile('SampleData.xlsx', {
            cellDates: true
        });
        const sheet_name_list = workbook.SheetNames;
        let excelData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);


        for (const userData of excelData) {


            const formatedDate = moment(userData.dob).format('YYYY-MM-DD').split('-');
            userData['bday'] = parseInt(formatedDate[2]) + 1;
            const bmonth = parseInt(formatedDate[1]);
            if (bmonth === 1) {
                userData['bmonth'] = 'January';
            } else if (bmonth === 2) {
                userData['bmonth'] = 'February';
            } else if (bmonth === 3) {
                userData['bmonth'] = 'March';
            } else if (bmonth === 4) {
                userData['bmonth'] = 'April';
            } else if (bmonth === 5) {
                userData['bmonth'] = 'May';
            } else if (bmonth === 6) {
                userData['bmonth'] = 'June';
            } else if (bmonth === 7) {
                userData['bmonth'] = 'July';
            } else if (bmonth === 8) {
                userData['bmonth'] = 'August';
            } else if (bmonth === 9) {
                userData['bmonth'] = 'September';
            } else if (bmonth === 10) {
                userData['bmonth'] = 'October';
            } else if (bmonth === 11) {
                userData['bmonth'] = 'November';
            } else if (bmonth === 12) {
                userData['bmonth'] = 'December';
            }

            userData['byear'] = formatedDate[0];

            //fillform function call
            await fillForm(userData);
        };

}


//fill the form
async function fillForm(user) {

    // get the driver
    let driver = await new Builder().forBrowser('chrome').build();

    //text field
    await driver.navigate().to('https://course.genesecloud.academy/login/signup.php')


    await driver.findElement(By.name("username")).sendKeys(user.username);
    await driver.findElement(By.name("password")).sendKeys(user.password);
    await driver.findElement(By.name("email")).sendKeys(user.email);
    await driver.findElement(By.name("email2")).sendKeys(user.email);
    await driver.findElement(By.name("firstname")).sendKeys(user.firstname);
    await driver.findElement(By.name("lastname")).sendKeys(user.lastname);
    await driver.findElement(By.name("city")).sendKeys(user.city);

    //dropdown
    await driver.findElement(By.name("country")).sendKeys("Nepal");
    await driver.findElement(By.name("profile_field_Province")).sendKeys(user.province);
    await driver.findElement(By.name("profile_field_citizenshipissueddistrict")).sendKeys(user.cid);
    await driver.findElement(By.name("profile_field_gender")).sendKeys(user.gender);
    // DOB
    await driver.findElement(By.name("profile_field_DOB[day]")).sendKeys(user.bday);
    await driver.findElement(By.name("profile_field_DOB[month]")).sendKeys(user.bmonth);
    await driver.findElement(By.name("profile_field_DOB[year]")).sendKeys(user.byear);

    //ok with default
    await driver.findElement(By.name("profile_field_Age")).sendKeys("21-30");
    await driver.findElement(By.name("profile_field_PWAD")).sendKeys("No");

    await driver.findElement(By.name("profile_field_Ethnic")).sendKeys(user.ethnic_group);
    await driver.findElement(By.name("profile_field_graduationyear")).sendKeys(user.gy);
    await driver.findElement(By.name("profile_field_Semester")).sendKeys(user.sem);

    await driver.findElement(By.name("profile_field_phoneM")).sendKeys(user.phone);
    await driver.findElement(By.name("profile_field_AddressC")).sendKeys(user.addressC);
    await driver.findElement(By.name("profile_field_AddressP")).sendKeys(user.addressP);
    await driver.findElement(By.name("profile_field_citizenship")).sendKeys(user.citizenship);
    await driver.findElement(By.name("profile_field_collegename")).sendKeys(user.col_name);
    await driver.findElement(By.name("profile_field_program")).sendKeys(user.program);
    await driver.findElement(By.name("profile_field_aws_educate_account")).sendKeys("No");
    await driver.findElement(By.name("profile_field_HCK")).sendKeys("Online");
    await driver.findElement(By.name("profile_field_ELocation")).sendKeys("Virtual");
    await driver.findElement(By.name("profile_field_Eventid")).sendKeys(user.eventid);


    // agree terms and condition
    await driver.findElement(By.name("policyagreed")).click();

    //submit form 
    await driver.findElement(By.name("submitbutton")).click();

    //let's wait for two seconds
    await driver.sleep(1000);

    // wait to get new title
    let title = await (await driver.getTitle()).trim();
    
    //close the browser if your title matches
    if(title == "Confirm your account") {
        await driver.close();
    }
      

}

// fillForm();