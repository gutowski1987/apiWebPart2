import Axios from "axios";
import styles from './DisplayPageWebPart.module.scss';
import { isNull } from "util";

window.url = window.location.href;
window.arrUrl = window.url.split('/');
window.starts;
window.arrUrl.forEach(function(e) {
    if(e.startsWith('Workspace-')) {
        window.starts = e;
    }
});
window.projectData = window.starts;

window.today = new Date();
var dd = String(today.getDate()).padStart(2, '0');
var mm = String(today.getMonth() + 1).padStart(2, '0'); //January is 0!
var yyyy = today.getFullYear();
window.today = dd + '/' + mm + '/' + yyyy;

window.projectId = window.projectData.split('-')[1];

$(window).ready(
    projectDetails(),
    businessCase(),
    risks(),
    budgetSummary(),
    financeFeasibility(),
    financeBudget(),
    financeAdditional(),
    projectTeam(),
    executiveStatus(),
    ragStatus(),
    projectRisks(),
    projectIssues(),
    completedThisPeriod(),
    focusForNextPeriod(),
    lessonsLearned(),
    gates(),
    keyMilestones()
);
function projectDetails() {
    $.ajax({
      type: 'GET',
      headers: {
        'accept': 'application/json;odata=verbose'
      },
      url: "/sites/portal/_api/web/lists(guid\'904b8e4d-93b3-41c3-b7b9-c36e6bc23f7c\')/items?ID,ProjectID,Title,Project_x0020_Status,ProjectType,Project_x0020_Category,Project_x0020_Category_x0020_Typ,Gate_x0020_Stage,Centre,ProjectCommencementDate,ExpectedDeliveryDate&$filter=ID eq \'" + projectId + "\'",
      success: 
    function(data){
          $(data.d.results).each(function(){
            $('#projectDetails').append("<h2>Project Summary - " + today + "</h2><h2>" + this.ID + " " + this.Title + "</h2><h3>Project Details</h3><table class=\"" + styles.WithoutBorder + "\"><thead><tr><th>Project Status</th><th>Project Type</th><th>Project Category</th><th>Project Category Type</th><th>Gate Stage</th><th>Centre</th><th>Project Commencement Date</th><th>Expected Delivery Date</th></tr></thead><tbody><tr><td>" + ifNull(this.Project_x0020_Status) + "</td><td>" + ifNull(this.ProjectType) + "</td><td>" + ifNull(this.Project_x0020_Category) + "</td><td>" + ifNull(this.Project_x0020_Category_x0020_Typ) + "</td><td>" + ifNull(this.Gate_x0020_Stage) + "</td><td>" + ifNull(this.Centre) + "</td><td>" + changeDate(this.ProjectCommencementDate) + "</td><td>" + changeDate(this.ExpectedDeliveryDate) + "</td></tr></tbody></table>");
        });
      }
    });
}
function businessCase() {
  $.ajax({
    type: 'GET',
    headers: {
      'accept': 'application/json;odata=verbose'
    },
    url: "/sites/portal/_api/web/lists(guid\'904b8e4d-93b3-41c3-b7b9-c36e6bc23f7c\')/items?BusinessCase&$filter=ID eq \'" + projectId + "\'",
    success: 
    function(data){
          $(data.d.results).each(function(){
            $('#businessCase').append("<h3>Business Case</h3><div>" + this.BusinessCase + "</div>");
        });
      }
  });
}
function risks() {
  $.ajax({
    type: 'GET',
    headers: {
      'accept': 'application/json;odata=verbose'
    },
    url: "/sites/portal/_api/web/lists(guid\'904b8e4d-93b3-41c3-b7b9-c36e6bc23f7c\')/items?RiskOfNotDoingTheProject&$filter=ID eq \'" + projectId + "\'",
    success: 
    function(data){
          $(data.d.results).each(function(){
            $('#risks').append("<h3>Risk of not doing the project</h3><div>" + ifNull(this.RiskOfNotDoingTheProject) + "</div>");
        });
      }
  });
}
function budgetSummary() {
  $.ajax({
    type: 'GET',
    headers: {
      'accept': 'application/json;odata=verbose'
    },
    url: "/sites/portal/_api/web/lists(guid\'904b8e4d-93b3-41c3-b7b9-c36e6bc23f7c\')/items?IncorporatedInBudget,WhereIsTheBudgetComingFrom,ForwardFunded_x0028_years_x0029_,BudgetYear&$filter=ID eq \'" + projectId + "\'",
    success: 
    function(data){
      $(data.d.results).each(function(){
        $('#budgetSummary').append("<h3>Budget Summary</h3><table><thead><tr><th>Incorporated into budget</th><th>Where is the budget coming from</th><th>Forwarded funded</th><th>Budget Year</th></tr></thead><tbody><tr><td>" + ifNull(this.IncorporatedInBudget) + "</td><td>" + ifNull(this.WhereIsTheBudgetComingFrom) + "</td><td>" + ifNull(this.ForwardFunded_x0028_years_x0029_) + "</td><td>" + ifNull(this.BudgetYear) + "</td></tr></tbody></table>");
    });
  }
  });
}
function financeFeasibility() {
  $.ajax({
    type: 'GET',
    headers: {
      'accept': 'application/json;odata=verbose'
    },
    url: "/sites/portal/_api/web/lists(guid\'904b8e4d-93b3-41c3-b7b9-c36e6bc23f7c\')/items?FeasibilityBudget,EstimatedBudget,EstimatedTotalBudget,FeasibilityOverallStatus&$filter=ID eq \'" + projectId + "\'",
    success: 
    function(data){
      $(data.d.results).each(function(){
        $('#financeFeasibility').append("<h3>Finance - Feasibility</h3><table><thead><tr><th>Feasibility Budget Requested</th><th>Estimated Budget</th><th>Estimated Total Budget</th><th>Feasibility Overall Status</th></tr></thead><tbody><tr><td>&#163; " + ifNull(this.FeasibilityBudget) + "</td><td>&#163; " + ifNull(this.EstimatedBudget) + "</td><td>&#163; " + ifNull(this.EstimatedTotalBudget) + "</td><td>" + ifNull(this.FeasibilityOverallStatus) + "</td></tr></tbody></table>");
    });
  }
  });
}
function financeBudget() {
  $.ajax({
    type: 'GET',
    headers: {
      'accept': 'application/json;odata=verbose'
    },
    url: "/sites/portal/_api/web/lists(guid\'904b8e4d-93b3-41c3-b7b9-c36e6bc23f7c\')/items?Budget_x0020_Overall_x0020_Statu,Budget,Actual_x0020_Budget_x0020_with_x&$filter=ID eq \'" + projectId + "\'",
    success: 
    function(data){
      $(data.d.results).each(function(){
        $('#financeBudget').append("<h3>Finance - Budget</h3><table><thead><tr><th>Actual Budget Requested</th><th>Total budget with feasibility</th><th>Budget Overall Status</th></tr></thead><tbody><tr><td>&#163; "+ ifNull(this.Budget) + "</td><td>&#163; " + ifNull(this.Actual_x0020_Budget_x0020_with_x) +"</td><td>" + ifNull(this.Budget_x0020_Overall_x0020_Statu) + "</td></tr></tbody></table>");
    });
  }
  });
}
function financeAdditional() {
  $.ajax({
    type: 'GET',
    headers: {
      'accept': 'application/json;odata=verbose'
    },
    url: "/sites/portal/_api/web/lists(guid\'63a6972f-902a-4102-86e4-1d06ff309a8c\')/items?FeasibilityBudget,MainBudget,AdditionalFinanceRequested,TotalOfALLBudgetswithALLAddition,OverallStatus&$filter=Title eq \'" + projectId + "\'",
    success: 
    function(data){
      $(data.d.results).each(function(){
        $('#financeAdditional').append("<h3>Finance - Additional</h3><table><thead><tr><th>Additional Finance Requested</th><th>Actual Budget</th><th>Total Budget with additional</th><th>Additional Budget Overall Status</th></tr></thead><tbody><tr><td>&#163; " + ifNull(this.AdditionalFinanceRequested) + "</td><td>&#163; " + ifNull(this.MainBudget) + "</td><td>&#163; " + ifNull(this.TotalOfALLBudgetswithALLAddition) + "</td><td>" + ifNull(this.OverallStatus) + "</td></tr></tbody></table>");
    });
  }
  });
}
function projectTeam() {
  $.ajax({
    type: 'GET',
    headers: {
      'accept': 'application/json;odata=verbose'
    },
    url: "/sites/portal/_api/web/lists(guid\'904b8e4d-93b3-41c3-b7b9-c36e6bc23f7c\')/items?ProjectManagerId,ProjectSponsorId,ProjectAdminId,GatekeeperOneId,GatekeeperTwoId,Stakeholders&$filter=ID eq \'" + projectId + "\'",
    success: 
    function(data){
      $(data.d.results).each(function(){
        $('#projectTeam').append("<h3>Project Team</h3><table><thead><tr><th>Project Manager</th><th>Project Sponsor</th><th>Project Admin</th><th>Gatekeeper One</th><th>Gatekeeper Two</th><th>Stakeholders</th></tr></thead><tbody><tr><td id=\"projectmanager\"></td><td id=\"projectsponsor\"></td><td id=\"projectadmin\"></td><td id=\"gatekeeperOne\"></td><td id=\"gatekeeperTwo\"></td><td id=\"stakeholder\"></td></tr></tbody></table>");
        callPerson(this.ProjectManagerId, 'projectmanager');
        callPerson(this.ProjectSponsorId, 'projectsponsor');
        callPerson(this.ProjectAdminId, 'projectadmin');
        callPerson(this.GatekeeperOneId, 'gatekeeperOne');
        callPerson(this.GatekeeperTwoId, 'gatekeeperTwo');
        callMultiPerson(this.StakeholdersId.results, 'stakeholder');
    });
  }
  });
}
function executiveStatus() {
  $.ajax({
    type: 'GET',
    headers: {
      'accept': 'application/json;odata=verbose'
    },
    url: "/sites/portal/_api/web/lists(guid\'659d7ed8-4e41-4498-b2a1-1ea932cc3c05\')/items?ExecutiveStatus,Created,User&$filter=ProjectID eq \'" + projectId + "\'&$top=1&$orderby=Created desc",
    success: 
    function(data){
        $(data.d.results).each(function(){
          $('#executiveStatus').append("<h3>Executive Status</h3><div>" + ifNull(this.ExecutiveStatus) + "</div><p><div>" + this.User + "</div><div>" + ifNull(changeDate(this.Created)) + "</div></p>");
      });
    }
  });
}
function ragStatus() {
  $.ajax({
      type: 'GET',
      headers: {
          'accept': 'application/json;odata=verbose'
      },
      url: "/sites/portal/_api/web/lists(guid\'2c61ef27-9e9d-4cf1-acc7-e301e54319f9\')/items?ProjectCostCode,ProjectCostIndicator,ProjectScopeCode,ProjectScopeIndicator,ProjectResourceCode,ProjectResourceIndicator,ProjectScheduleCode,ProjectScheduleInducator,Created&$filter=Project_x0020_ID eq \'" + projectId + "\'&$orderby=Created desc&$top=1",
      success: 
      function(data){
          $(data.d.results).each(function(){
              icon(this.ProjectCostIndicator, '#iconCost');
              icon(this.ProjectResourceIndicator, '#iconResource');
              icon(this.ProjectScopeIndicator, '#iconScope');
              icon(this.ProjectScheduleIndicator, '#iconSchedule');
              color(this.ProjectCostCode, '#colorCost');
              color(this.ProjectResourceCode, '#colorResource');
              color(this.ProjectScopeCode, '#colorScope');
              color(this.ProjectScheduleCode, '#colorSchedule');
              $('#ragDate').append(changeDate(this.Created));
          });
      }
  });
}
function projectRisks() {
  $.ajax({
    type: 'GET',
    headers: {
      'accept': 'application/json;odata=verbose'
    },
    url: "/sites/portal/_api/web/lists(guid\'0245e024-43ac-4cae-9f69-75f23392b803\')/items?RiskReference,RiskDescription,Gross_x0020_Impact,Risk_x0020_Mitigation,Net_x0020_Impact,Risk_x0020_OwnerId&$filter=ProjectID eq \'" + projectId + "\'",
    success: 
  function(data){
    var personCounter = 1;
        $(data.d.results).each(function(){
          $('#projectRisks').append("<tr><td>" + ifNull(this.RiskReference) + "</td><td>" + ifNull(this.RiskDescription) + "</td><td class=\"" + styles.narrow + "\">" + ifNull(this.Gross_x0020_Impact) + "</td><td>" + ifNull(this.Risk_x0020_Mitigation) + "</td><td class=\"" + styles.narrow + "\">" + ifNull(this.Net_x0020_Impact) + "</td><td id=\"riskOwner" + personCounter + "\"></td></tr>");
      callPerson(this.Risk_x0020_OwnerId, "riskOwner" + personCounter);
      personCounter++;
      });
    }
  });
}
function projectIssues() {
  $.ajax({
    type: 'GET',
    headers: {
      'accept': 'application/json;odata=verbose'
    },
    url: "/sites/portal/_api/web/lists(guid\'b3237ca3-e8e8-4af5-a070-e664aa57fdd0\')/items?RiskReference,IssueDescription,Escalation_x0020_Level,Actions_x0020_to_x0020_Resolve,Actions_x0020_Due,Action_x0020_OwnerId&$filter=ProjectID eq \'" + projectId + "\'",
    success: 
      function(data){
        var personCounter = 1;
        $(data.d.results).each(function(){
              $('#projectIssues').append("<tr><td>" + ifNull(this.RiskReference) + "</td><td>" + ifNull(this.IssueDescription) + "</td><td>" + ifNull(this.Escalation_x0020_Level) + "</td><td>" + ifNull(this.Actions_x0020_to_x0020_Resolve) + "</td><td>" + changeDate(this.Actions_x0020_Due) + "</td><td    id=\"actionOwner" + personCounter + "\"></td></tr>");
          callPerson(this.Action_x0020_OwnerId, "actionOwner" + personCounter);
          personCounter++;
        });
      }
    });
}
function completedThisPeriod() {
  $.ajax({
    type: 'GET',
    headers: {
      'accept': 'application/json;odata=verbose'
    },
    url: "/sites/portal/_api/web/lists(guid\'6c56bc85-9f8b-498b-acea-94ba5d400642\')/items?Task,Date,OwnerId&$filter=(Title eq \'" + projectId + "\') and (Catagory eq \'Completed This Period\')&$orderby=Date desc",
    success: 
  function(data){
    var personCounter = 1;
        $(data.d.results).each(function(){
          $('#completedThisPeriod').append("<tr><td>" + ifNull(this.Task)+ "</td><td class=\"" + styles.narrow + "\">" + changeDate(this.Date)+ "</td><td id=\"periodOwner" + personCounter + "\"></td></tr>");
          callPerson(this.OwnerId, "periodOwner" + personCounter);
          personCounter++;
      });
    }
  });
}
function focusForNextPeriod() {
  $.ajax({
    type: 'GET',
    headers: {
      'accept': 'application/json;odata=verbose'
    },
    url: "/sites/portal/_api/web/lists(guid\'6c56bc85-9f8b-498b-acea-94ba5d400642\')/items?Task,Date,OwnerId&$filter=(Title eq \'" + projectId + "\') and (Catagory eq \'Focus For Next Period\')&$orderby=Date desc",
    success: 
    function(data){
      var personCounter = 1;
          $(data.d.results).each(function(){
            $('#focusForNextPeriod').append("<tr><td>" + ifNull(this.Task)+ "</td><td class=\"" + styles.narrow + "\">" + changeDate(this.Date)+ "</td><td id=\"focusOwner" + personCounter + "\"></td></tr>");
            callPerson(this.OwnerId, "focusOwner" + personCounter);
            personCounter++;
        });
      }
  });
}
function gates() {
  $.ajax({
    type: 'GET',
    headers: {
      'accept': 'application/json;odata=verbose'
    },
    url: "/sites/portal/_api/web/lists(guid\'3e8b7736-85b8-465f-a93d-da859a5639a4\')/items?CurrentGate,RequestedGate,Comment,Outcome,Created,Date&$filter=ProjectID eq \'" + projectId + "\'",
    success: 
  function(data){
        $(data.d.results).each(function(){
          $('#gates').append("<tr><td>" + ifNull(this.CurrentGate) + "</td><td>" + ifNull(this.RequestedGate) + "</td><td>" + ifNull(this.Comment) + "</td><td>" + ifNull(this.Outcome) + "</td><td class=\"" + styles.narrow + "\">" + changeDate(this.Created) + "</td><td class=\"" + styles.narrow + "\">" + ifNull(this.Date) + "</td></tr>");
      });
    }
  });
}
function lessonsLearned() {
  $.ajax({
    type: 'GET',
    headers: {
      'accept': 'application/json;odata=verbose'
    },
    url: "/sites/portal/_api/web/lists(guid\'e5b8658d-5c63-4bc9-babc-022c2062b81a\')/items?SituationDescription,LessonsLearnt,FollowUpAction,ActionDate,ActionedById,Outcome&$filter=(ProjectID eq \'" + projectId + "\')",
    success: 
  function(data){
    var personCounter = 1;
        $(data.d.results).each(function(){
          $('#lessonsLearned').append("<tr><td>" + ifNull(this.SituationDescription) + "</td><td>" + ifNull(this.LessonsLearnt) + "</td><td>" + ifNull(this.FollowUpAction) + "</td><td class=\"" + styles.narrow + "\">" + changeDate(this.ActionDate) + "</td><td id=\"lessonOwner" + personCounter + "\"></td><td>" + ifNull(this.Outcome) + "</td></tr>");
          callPerson(this.ActionedById, "lessonOwner" + personCounter);
          personCounter++;
      });
    }
  });
}
function keyMilestones() {
  $.ajax({
    type: 'GET',
    headers: {
      'accept': 'application/json;odata=verbose'
    },
    url: "/sites/portal/_api/web/lists(guid\'0a170a49-6e30-481f-9094-d75d511b904c\')/items?MilestoneType,Date,MOrder,&$filter=ProjectID eq \'" + projectId + "\'&$top=12&$orderby=MOrder asc",
    success: 
    function(data){
        $(data.d.results).each(function(){
          $('#keyMilestones').append("<tr><td>"+ ifNull(this.MOrder) + " - " + ifNull(this.MilestoneType) + "</td><td><div>" + ifNull(this.Gate) + "</div></td><td class=\"" + styles.narrow + "\">" + changeDate(this.Date) + "</td></tr>");
      });
    }
  });
}

// helpers
function color(data, where) { 
  if(data === 'Green') {
      $(where).addClass("ms-bgColor-greenLight");
  } else if(data === 'Amber') {
      $(where).addClass("ms-bgColor-yellow");
  } else if(data === 'Red') {
      $(where).addClass("ms-bgColor-red");
  }
}
function icon(data, where) {
  if(data === '=') {
      $(where).append("<span style='font-size:25px;'>&#61;</span>");
  } else if(data === '<') {
      $(where).append("<span style='font-size:25px;'>&#8600;</span>");
  } else if(data === '>') {
      $(where).append("<span style='font-size:25px;'>&#8599;</span>");
  }
}
function callPerson(identificator, where) {
  var resultElement = document.getElementById(where);
  resultElement.innerHTML = '';
  if(identificator == null) {
    return
  } else {
    Axios.get("https://intu.sharepoint.com/sites/portal/_api/web/getuserbyid(" + identificator + ")")
    .then(function (response) {
      resultElement.innerHTML = response.data.Title;
    })  
  }
}
function callMultiPerson(identificators, where) {
  if(identificators == null) {
    return " ";
  } else {

    for (var i = 0; i <  Object.entries(identificators).length; i++) {
      Axios.get("https://intu.sharepoint.com/sites/portal/_api/web/getuserbyid(" + Object.entries(identificators)[i][1] + ")")
      .then(function (response) {
        $("#" + where).append("<div>" + response.data.Title + "</div>");
      })  
    }
  }
}
function changeDate(input) {
if(input == null) {
return " ";
} else {
var stringDate = input.slice(0,10).split("-");
var readyDate = stringDate[2] + "-" + stringDate[1] + "-" + stringDate[0];
return readyDate;
}
}
function ifNull(val) {
if(val == null) {
return " ";
} else {
return val;
}
}