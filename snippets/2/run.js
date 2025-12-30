function testFunction(name) {
    return `Hello, ${name}!`;
}

const myname = Xrm.Utility.getGlobalContext().userSettings.userName;
const greeting = testFunction(myname);
window.alert(greeting);
