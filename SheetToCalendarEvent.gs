/**
 * This will create a custom menu in your Google Sheet everytime you open the Sheet
 */
function onOpen(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Apps')
    .addItem('Add People to Event','addAttendeesToEvent')
    .addToUi()
}


function addAttendeesToEvent(){
  const calendar = CalendarApp.getDefaultCalendar(); 
  const ui = SpreadsheetApp.getUi();

  const eventID = getTextFromPrompt('Enter the eventID of the Calendar Event you want to add attendees to:');
  const event = calendar.getEventById(eventID); 

  if(event !== null){

    event.setGuestsCanSeeGuests(false);
    const guests = getGuests();

    const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
    let validGuests = 0;

    guests.forEach(guest => {
      let email = guest.trim();

      try {
        if (emailRegex.test(email)) {
          event.addGuest(email);
          validGuests++;
        } else {
          console.warn(`${email} is not a valid Email!`);
        }
      } catch (err) {
        console.error(err);
      }
    });

    ui.alert('✅ Success',`${validGuests} out of ${guests.length} ${guests.length == 1 ?'person': 'people'} ${guests.length == 1 ?'was': 'were'} added to the calendar event`,ui.ButtonSet.OK);
  } else{
    ui.alert('🚨 Error','Not a valid EventID',ui.ButtonSet.OK);
  }
}

/**
 * This will return an array of emails that are in a specified column
 * @returns {string[]} array of emails
 */
function getGuests(){
  const activeSheet = SpreadsheetApp.getActiveSheet();
  const sheetData = activeSheet.getDataRange().getValues();
  sheetData.shift();

  const emailsColumn = getTextFromPrompt(
    'Enter the Column number that contains the emails of the attendees.'
      + ' Column count starts from 0, so col A is 0, B is 1 etc...:'
    );

  return sheetData.map(item => item[emailsColumn])
}

function getTextFromPrompt(question) {
  const ui = SpreadsheetApp.getUi(); 

  const response = ui.prompt(question);
  return response.getResponseText();
}
