const ACCOUNT_SHEET = 'Account'
const accountSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACCOUNT_SHEET)

const INSTANCE_URL = PropertiesService.getScriptProperties().getProperty("INSTANCE_URL")
const QUERY_URL = PropertiesService.getScriptProperties().getProperty("QUERY_URL")

const AUTH_TOKEN_URL=PropertiesService.getScriptProperties().getProperty("AUTH_TOKEN_URL")

const _getAuthToken = () => {

  try{
    response = UrlFetchApp.fetch(AUTH_TOKEN_URL)
  }catch(e){
    SpreadsheetApp.getUi().alert("ERROR " + e)
  }

  return JSON.parse(response).access_token
}

const _validateInput = (input) => {
  
  let regex = /^[A-Za-z]+$/
  if(input.match(regex)){
    return true
  }
  return false
}

const config = {
        headers: {
            "Authorization": "Bearer " + _getAuthToken()
          }
        }


// Main Account Presentation Layer Offsets
const columnOffset = 5
const rowOffset = 3
const checkboxOffset = 8
const accountsRange = accountSheet.getRange(3, 5, 10, 4)
const checkboxRange = accountSheet.getRange(8, 5, 10, 1)

// Find Accounts 
const onFindAccounts = () => {

  let accountName = accountSheet.getRange(3, 3).getValue()
  let accountOwner = accountSheet.getRange(4, 3).getValue()

  // validate accountName and accountOwner
  if(!_validateInput(accountName)){
    SpreadsheetApp.getUi().alert("Account Name must contain letters only no special characters or spaces")
    return  
  }

  if(!_validateInput(accountOwner)){
    SpreadsheetApp.getUi().alert("Account Owner must contain letters only no special characters or spaces")
    return  
  }

  let query = "select+Account.Id,+Account.Name,+Account.Account_Owner_Full_Name__c,+Account.Website+from+Account+where+Account_Owner_Full_Name__c+like+'%25" + accountOwner + "%25'+AND+Account.Name+like+'%25" + accountName +"%25'+limit+10"
  let response = ""

  // clear the accounts data range
  accountsRange.clear()
  accountsRange.removeCheckboxes()

  try{

    response = UrlFetchApp.fetch(INSTANCE_URL+QUERY_URL+query, config)

    if(JSON.parse(response).totalSize == 0){
      SpreadsheetApp.getUi().alert("There is no data please refine your search")
    }else{

      for(i=3; i < (JSON.parse(response).totalSize + 3); i++){      
        accountSheet.getRange(i,5).setValue(JSON.parse(response).records[i-3].Id)
        accountSheet.getRange(i,6).setValue(JSON.parse(response).records[i-3].Name)
        accountSheet.getRange(i,7).setValue(JSON.parse(response).records[i-3].Account_Owner_Full_Name__c)
        accountSheet.getRange(i,8).insertCheckboxes()
      }

    }
  }catch(e){
    console.log(e)
    SpreadsheetApp.getUi().alert(e)
  }

}

const createResearchTemplate = () => {

  let account = {}

  // check that a checkbox has been checked
  for(i=0; i<10; i++){
      if(accountSheet.getRange(i+rowOffset, checkboxOffset).getValue() == true){
        account.id = accountSheet.getRange(i+rowOffset, columnOffset).getValue()
        account.name = accountSheet.getRange(i+rowOffset, columnOffset+1).getValue()
        account.owner = accountSheet.getRange(i+rowOffset, columnOffset+2).getValue()
      }
  }
  if(account.id == null || account.id == undefined){
    SpreadsheetApp.getUi().alert("Please select one account from the list")
  }else{
    createDocument(account)
  }

}

const createDocument = (account) => {

  const doc = DocumentApp.create(account.name)
  const body = doc.getBody()

  // setup Header and Account Details
  let header = body.appendParagraph(account.name + " Customer Research")
  header.setHeading(DocumentApp.ParagraphHeading.TITLE)
  header.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
  
  // Account Team
  let accountTeam = body.appendParagraph("Account Team")
  accountTeam.setHeading(DocumentApp.ParagraphHeading.HEADING1)
  let accountTeamDetails = []
  accountTeamDetails.push(["Account Owner", account.owner])
  accountTeamDetails.push(["Account Owner Manager", " "])
  accountTeamDetails.push(["Solutions Engineer", " "])
  accountTeamDetails.push(["Technical Account Manager", " "])
  body.appendTable(accountTeamDetails)

  // Customer Details
  // let customerDetails = body.appendParagraph("Customer Details")
  // customerDetails.setHeading(DocumentApp.ParagraphHeading.HEADING1)

  // // Financials
  // let financials = body.appendParagraph("Financials")
  // financials.setHeading(DocumentApp.ParagraphHeading.HEADING1)

  // get contacts
  let contactsQuery = "select+Contact.CreatedDate,+Contact.Name,+Contact.Title,+Contact.Contact_Status__c+from+Contact+where+AccountId+=+'" + account.id +"'+and+Contact.CreatedDate!=NULL+and+Contact.Name+!=+NULL+and+Contact.Title+!=+NULL+and+Contact.Contact_Status__c+IN+(+'Developer',+'Opportunity',+'Open',+'Connected',+'Added',+'Engaged',+'Nurture'+)+order+by+Contact.CreatedDate+desc+limit+100"
  let contacts = []
  
  // Set header for table
  contacts.push(["Created Date", "Name", "Status"])

  try{

    let response = UrlFetchApp.fetch(INSTANCE_URL+QUERY_URL+contactsQuery, config)
    let data = JSON.parse(response)

    // parse the data
    for(i=0; i<data.totalSize; i++){
      let createdDate = new Date(data.records[i].CreatedDate).toLocaleDateString()
      let contactName = data.records[i].Name
      // let contactTitle = data.records[i].Title
      let contactStatus = data.records[i].Contact_Status__c
      contacts.push([createdDate, contactName, contactStatus])
    }

    if(contacts.length > 1){
      let contactsBody = body.appendParagraph("Contacts")
      contactsBody.setHeading(DocumentApp.ParagraphHeading.HEADING1)
      body.appendTable(contacts)
    }

  }catch(e){
    console.log(e)
    SpreadsheetApp.getUi().alert(e)
  }


  // get support tickets
  let supportTicketsQuery = "select+Customer_Support_Ticket__c.Ticket_ID__c,+Customer_Support_Ticket__c.CreatedDate,+Customer_Support_Ticket__c.Date_Solved__c,+Customer_Support_Ticket__c.Priority__c,+Customer_Support_Ticket__c.Ticket_Subject__c+from+Customer_Support_Ticket__c+where+Account__c='" + account.id + "'+and+Customer_Support_Ticket__c.Ticket_ID__c!=NULL+and+Customer_Support_Ticket__c.CreatedDate!=NULL+and+Customer_Support_Ticket__c.Date_Solved__c!=NULL+and+Customer_Support_Ticket__c.Priority__c!=NULL+order+by+Customer_Support_Ticket__c.CreatedDate+desc+limit+100"
  

  // +and+Customer_Support_Ticket__c.Ticket_Subject__c!=NULL
  // +order+by+Customer_Support_Ticket__c.CreatedDate+desc+limit+100"
// +And+Customer_Support_Ticket__c!=’SCRUBBED’+order+by+Customer_Support_Ticket__c.CreatedDate+desc+limit+100

  let supportTickets = []

  try{

    let response = UrlFetchApp.fetch(INSTANCE_URL+QUERY_URL+supportTicketsQuery, config)
    let data = JSON.parse(response)

    // set the header
    supportTickets.push(["Created Date", "Date Solved", "Ticket Id", "Subject"])

    // parse the data
    for(i=0; i<data.totalSize; i++){
      let ticketId = data.records[i].Ticket_ID__c
      let createdDate = new Date(data.records[i].CreatedDate)
      let dateSolved = new Date(data.records[i].Date_Solved__c)
      // let priorty = data.records[i].Priority__c
      let subject = data.records[i].Ticket_Subject__c

      let now = new Date()

      // we only want tickets within a certain range
      if((now.getFullYear() - createdDate.getFullYear()) < 2){
        supportTickets.push([createdDate.toLocaleDateString(), dateSolved.toLocaleDateString(), ticketId, subject])
      }
      
    }

    if(supportTickets.length > 1){
      let supportBody = body.appendParagraph("Support Tickets")
      supportBody.setHeading(DocumentApp.ParagraphHeading.HEADING1)
      body.appendTable(supportTickets)
    }

  }catch(e){
    console.log(e)
    SpreadsheetApp.getUi().alert(e)
  }

  // get opportunities 
  let opportunitiesQuery = "select+Opportunity.StageName,+Opportunity.Opportunity_Name_First_80_Characters__c,+Opportunity.eARR_post_Launch__c,+Opportunity.CloseDate+from+Opportunity+where+AccountId='" + account.id + "'+and+Opportunity.StageName!=NULL+and+Opportunity.Opportunity_Name_First_80_Characters__c!=NULL+and+Opportunity.eARR_post_Launch__c!=NULL+and+Opportunity.CloseDate!=NULL+order+by+Opportunity.CloseDate+desc+limit+100"

  let lostOpportunities = []
  let wonOpportunities = []
  let openOpportunities = []

  // Set the headers of the table
  lostOpportunities.push(["Stage Name", "Name", "eARR", "Close Date"])
  wonOpportunities.push(["Stage Name", "Name", "eARR", "Close Date"])
  openOpportunities.push(["Stage Name", "Name", "eARR", "Close Date"])

  try{

    let response = UrlFetchApp.fetch(INSTANCE_URL+QUERY_URL+opportunitiesQuery, config)
    let data = JSON.parse(response)

    for(i=0; i<data.totalSize; i++){
      let stageName = data.records[i].StageName
      let opportunityName = data.records[i].Opportunity_Name_First_80_Characters__c
      let earr = data.records[i].eARR_post_Launch__c
      let closeDate = new Date(data.records[i].CloseDate).toLocaleDateString()

      if(stageName == "Closed Lost"){
        lostOpportunities.push([stageName, opportunityName, earr, closeDate])
      }
      if(stageName == "Closed Won"){
        wonOpportunities.push([stageName, opportunityName, earr, closeDate])
      }
      if(stageName == "Scope" || stageName == "Qualified" || stageName == "Validate Solution" || stageName == "Submit Proposal" || stageName == "Commit"){
        openOpportunities.push([stageName, opportunityName, earr, closeDate])
      }

    }

    if(lostOpportunities.length > 1){
      let opportunityBody = body.appendParagraph("Lost Opportunities")
      opportunityBody.setHeading(DocumentApp.ParagraphHeading.HEADING1)
      body.appendTable(lostOpportunities)
    }
    if(wonOpportunities.length > 1){
      let opportunityBody = body.appendParagraph("Won Opportunities")
      opportunityBody.setHeading(DocumentApp.ParagraphHeading.HEADING1)
      body.appendTable(wonOpportunities)
    }
    if(openOpportunities.length > 1){
      let opportunityBody = body.appendParagraph("Open Opportunities")
      opportunityBody.setHeading(DocumentApp.ParagraphHeading.HEADING1)
      body.appendTable(openOpportunities)
    }
    
  }catch(e){
    console.log(e)
    SpreadsheetApp.getUi().alert(e)
  }

  // Better Business Bureau
  let bbb = body.appendParagraph("Better Business Bureau")
  bbb.setHeading(DocumentApp.ParagraphHeading.HEADING1)
  let subBBB = body.appendParagraph("What is their Better Business Bureau rating and any associated reviews which may highlight any customer experience issues")
  subBBB.setHeading(DocumentApp.ParagraphHeading.NORMAL)
  let bbbSearch = body.appendParagraph("\nGoogle Search Link")
  bbbSearch.setHeading(DocumentApp.ParagraphHeading.NORMAL)
  bbbSearch.setLinkUrl("https://www.google.com/search?q=bbb+" + account.name)
  
  // Glassdoor
  let glassdoor = body.appendParagraph("Glassdoor")
  glassdoor.setHeading(DocumentApp.ParagraphHeading.HEADING1)
  let subGlassdoor = body.appendParagraph("Do the employees mention any broken business processes? Is there any volatility in the company?")
  subGlassdoor.setHeading(DocumentApp.ParagraphHeading.NORMAL)
  let glassdoorSearch = body.appendParagraph("\nGoogle Search Link")
  glassdoorSearch.setHeading(DocumentApp.ParagraphHeading.NORMAL)
  glassdoorSearch.setLinkUrl("https://www.google.com/search?q=glassdoor+" + account.name)

  // News
  let news = body.appendParagraph("News")
  news.setHeading(DocumentApp.ParagraphHeading.HEADING1)
  let subNews = body.appendParagraph("What has been in the news lately? any rounds of investment? Any layoffs? Anything of note?")
  subNews.setHeading(DocumentApp.ParagraphHeading.NORMAL)

  // Customer Experience
  let customerExperience = body.appendParagraph("Customer Experience")
  customerExperience.setHeading(DocumentApp.ParagraphHeading.HEADING1)
  let subCustomerExperience = body.appendParagraph("This is an opportunity to do some mystery shopping. Does the customer have a public facing contact channel like chat, phone etc? Do they have an IVR tree, or an app? What is the customer experience like?")
  subCustomerExperience.setHeading(DocumentApp.ParagraphHeading.NORMAL)

  // Similar Customers
  let similarCustomers = body.appendParagraph("Similar Customers")
  similarCustomers.setHeading(DocumentApp.ParagraphHeading.HEADING1)
  let subSimilarCustomers = body.appendParagraph("Are there similar customers that use Twilio? What type of use cases have they deployed?")
  subSimilarCustomers.setHeading(DocumentApp.ParagraphHeading.NORMAL)

  // Technology
  let technology = body.appendParagraph("Technology")
  technology.setHeading(DocumentApp.ParagraphHeading.HEADING1)
  let subTechnology = body.appendParagraph("What is their technology stack?")
  subTechnology.setHeading(DocumentApp.ParagraphHeading.NORMAL)
  let twilio = body.appendParagraph("Twilio:")
  twilio.setHeading(DocumentApp.ParagraphHeading.HEADING2)
  let zoominfo = body.appendParagraph("Zoominfo:")
  zoominfo.setHeading(DocumentApp.ParagraphHeading.HEADING2)

  // Analysis
  let analysis = body.appendParagraph("Analysis")
  analysis.setHeading(DocumentApp.ParagraphHeading.HEADING1)

  // Hypothesis
  let hypothesis = body.appendParagraph("Hypothesis")
  hypothesis.setHeading(DocumentApp.ParagraphHeading.HEADING2)
  let subHypothesis = body.appendParagraph("Based on what you’ve learned, what business problems do you believe we can help with?  What solutions and/or use cases should we disco around for them?")
  subHypothesis.setHeading(DocumentApp.ParagraphHeading.NORMAL)

  // Sales Play
  let salesPlay = body.appendParagraph("Sales Play")
  salesPlay.setHeading(DocumentApp.ParagraphHeading.HEADING2)
  let subSalesPlay = body.appendParagraph("e.g. Delivery Notifications, Account Security, Intelligent Chatbot etc")
  subSalesPlay.setHeading(DocumentApp.ParagraphHeading.NORMAL)

  //Store the url of our new document in a variable
  const url = doc.getUrl()

  // Display the output
  const ui = SpreadsheetApp.getUi()
  const htmlString = "<base target=\"_blank\"><a href=\"" + url + "\">" + account.name + "</a>"
  const html = HtmlService.createHtmlOutput(htmlString)
  ui.showModalDialog(html, 'Document Created')



}
