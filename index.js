
let API = 'https://graph.microsoft.com/v1.0//me/findMeetingTimes'
let all_event_request = 'https://graph.microsoft.com/v1.0/me/events?$select=subject,body,bodyPreview,organizer,attendees,start,end,location'

let TOKEN = 'eyJ0eXAiOiJKV1QiLCJub25jZSI6IklBV0tfTGpaM1hUTnZKcGVEMDVFMGM3b3VWMDh0RVZMbG1ROHdsWS0yc2MiLCJhbGciOiJSUzI1NiIsIng1dCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyIsImtpZCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC8yYTE5MDdjMy01OTJjLTQwY2ItOGE4OC03ZjJkNDBiZDkzN2EvIiwiaWF0IjoxNjcyMTcyMjMxLCJuYmYiOjE2NzIxNzIyMzEsImV4cCI6MTY3MjE3NzUwOSwiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFWUUFxLzhUQUFBQVdCWmJlYk5PdDUxcG56em5jOGxuaHFnNGl3Q1NBRFRUN3kvdzI1Z3hLNytWSUJvZ1F0QXUxUXlPSVVoY25XMzE5QllPMGtETzlmVk56TlNxOGxTdlRNYzBNVllQS3pzZWZ2dFpQRy9pV3JjPSIsImFtciI6WyJwd2QiLCJtZmEiXSwiYXBwX2Rpc3BsYXluYW1lIjoiT3V0bG9va0Jvb2tpbmciLCJhcHBpZCI6ImU3MDI4NDNhLTY1MzEtNGQ0Mi05Y2M0LTQ4OGUzMzRjMjc3OSIsImFwcGlkYWNyIjoiMCIsImZhbWlseV9uYW1lIjoiUm9kcmlndWVzIiwiZ2l2ZW5fbmFtZSI6IkFydGVyaW8iLCJpZHR5cCI6InVzZXIiLCJpcGFkZHIiOiIxNDYuOTUuMTgwLjE1NyIsIm5hbWUiOiJBcnRlcmlvIFJvZHJpZ3VlcyIsIm9pZCI6IjUxNzNhOGM4LWY4YTEtNDMyZS1hYzk3LWUyNTFlMmU2MDJmMSIsIm9ucHJlbV9zaWQiOiJTLTEtNS0yMS0yNzY0Nzc4OTQ3LTM5MjgzMTM4NTYtMzIzNzA1MjM1Ni0xMTY3MTQiLCJwbGF0ZiI6IjUiLCJwdWlkIjoiMTAwMzIwMDAzNkYwRDI1QyIsInJoIjoiMC5BVGNBd3djWktpeFp5MENLaUg4dFFMMlRlZ01BQUFBQUFBQUF3QUFBQUFBQUFBQTNBUFkuIiwic2NwIjoiTWFpbC5SZWFkIG9wZW5pZCBwcm9maWxlIFVzZXIuUmVhZCBlbWFpbCIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6InJyZDEyOV8xUVE2RkZUMkxJVEs5YWFqc1VUWVFJcmF0UnlzTk96SXZPWlUiLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiTkEiLCJ0aWQiOiIyYTE5MDdjMy01OTJjLTQwY2ItOGE4OC03ZjJkNDBiZDkzN2EiLCJ1bmlxdWVfbmFtZSI6IkFydGVyaW8uUm9kcmlndWVzNDdAbXlodW50ZXIuY3VueS5lZHUiLCJ1cG4iOiJBcnRlcmlvLlJvZHJpZ3VlczQ3QG15aHVudGVyLmN1bnkuZWR1IiwidXRpIjoiWFd2cm9ZLUJqRU9vYng2Y3M5cVNBUSIsInZlciI6IjEuMCIsIndpZHMiOlsiYjc5ZmJmNGQtM2VmOS00Njg5LTgxNDMtNzZiMTk0ZTg1NTA5Il0sInhtc19zdCI6eyJzdWIiOiJuUUhHdHdEaHBGN1NwM1NwZ0NIeUoxWmR2anhnWlF0YWhvVUJnb1dVcVBRIn0sInhtc190Y2R0IjoxMzgyMDM1MjU5fQ.fFgrw8slMvW-zFw92a91JqJISLcnDI1PYrc7j9Jy-b6Z42Ov8rEL9wN2MGhA9l4-HZWRZf2o0ZLpL74DlsR9TBImqiNyzRDfCZ2gFwbsxoQSp60fJtwquW_p4vgMcImQ5TyFfG5nwb_T-Y2rj_rffCrHkYIylIYV9h6irYWTY68yzX1JEnFBog4bGeM-OViYDb4Zvwvq_FmJXZjVOEii23d0Iiy5lbd36VLzAnbi4vnXaRD5QMIeVxYiQzS18Ex-31HjQ9fmoDO5rLbPB4R6EYLuZRW42el9PPPbHji7j1Z1L3SFtoq0NWfWNzpHUEYvlN5hv2WFWPdOob3H6KlJ1A'

let ENDPOINT_Calendar = "https://graph.microsoft.com/v1.0/me/calendar"
let ENDPOINT_Message = "https://graph.microsoft.com/v1.0/me/messages"
let ENDPOINT = "https://graph.microsoft.com/v1.0/me/"
let ENDPOINT_FindMeetingTimes = "https://graph.microsoft.com/v1.0/me/findMeetingTimes"

let CALLBACK = (res, endpoint) => {
    let doc = document.getElementById('AllAppointment');

    console.log(res)

    for (const property in res){
        doc.append(`${property}, ${res[property]}`);
    }


}

const meetingTimeSuggestionsResult = {
    attendees: [ 
      { 
        type: 'required',  
        emailAddress: { 
          name: 'Alex Wilbur',
          address: 'alexw@contoso.onmicrosoft.com' 
        } 
      }
    ],  
    locationConstraint: { 
      isRequired: false,  
      suggestLocation: false,  
      locations: [ 
        { 
          resolveAvailability: false,
          displayName: 'Conf room Hood' 
        } 
      ] 
    },  
    timeConstraint: {
      activityDomain: 'work', 
      timeSlots: [ 
        { 
          start: { 
            dateTime: '2019-04-16T09:00:00',  
            timeZone: 'Pacific Standard Time' 
          },  
          end: { 
            dateTime: '2019-04-18T17:00:00',  
            timeZone: 'Pacific Standard Time' 
          } 
        } 
      ] 
    },  
    isOrganizerOptional: 'false',
    meetingDuration: 'PT1H',
    returnSuggestionReasons: 'true',
    minimumAttendeePercentage: '100'
};

function callMSGraph(endpoint, token, callback) {
    const headers = new Headers();
    const bearer = `Bearer ${token}`;

    headers.append("Authorization", bearer);

    const options = {
        method: "POST",
        headers: headers,
        authProvider: token,
    };

    console.log('request made to Graph API at: ' + new Date().toString());

    fetch(endpoint, options, meetingTimeSuggestionsResult)
        .then(response => response.json())
        .then(response => callback(response, endpoint))
        .catch(error => console.log(error));
}

window.onload = async () => {
    callMSGraph(ENDPOINT_FindMeetingTimes, TOKEN, CALLBACK)
}
