import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import styles from './DisplayPageWebPart.module.scss';
import * as strings from 'DisplayPageWebPartStrings';

import * as $ from 'jquery';
import axios from 'axios';

export interface IDisplayPageWebPartProps {
  description: string;
}

export default class DisplayPageWebPart extends BaseClientSideWebPart<IDisplayPageWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class="ms-Fabric" dir="ltr">
      <div id="projectDetails"></div>
    
      <div id="businessCase" class="ProjectObjective"></div>

      <div id="risks" class="ExecutiveStatus"></div>

      <div id="budgetSummary"></div>

      <div id="financeFeasibility"></div>

      <div id="financeBudget"></div>

      <div id="financeAdditional"></div>

      <div id="projectTeam"></div>

      <div id="executiveStatus"></div>

      <div class="${styles.RAGStatus}">
          <h3>RAG Status</h3>
          <div id="ragDate"></div>
          <table>
          <thead>
              <tr>
                  <th class="ms-textAlignCenter">Project Cost</th>
                  <th class="ms-textAlignCenter">Project Scope</th>
                  <th class="ms-textAlignCenter">Project Resource</th>
                  <th class="ms-textAlignCenter">Project Schedule</th>
              </tr>
          </thead>
          <tbody>
              <tr>
                  <td id="colorCost"></td>
                  <td id="colorScope"></td>
                  <td id="colorResource"></td>
                  <td id="colorSchedule"></td>
              </tr>
              <tr>
                  <td id="iconCost" class="ms-textAlignCenter"></td>
                  <td id="iconScope" class="ms-textAlignCenter"></td>
                  <td id="iconResource" class="ms-textAlignCenter"></td>
                  <td id="iconSchedule" class="ms-textAlignCenter"></td>
              </tr>
          </tbody>
          </table>
      </div>

      <div class="projectRisk">
      <h3>Project Risks</h3>
      <table>
        <thead>
          <tr>
            <th>Risk Ref</th>
            <th>Description</th>
            <th class="narrow">Gross Risk</th>
            <th>Mitigation</th>
            <th class="narrow">Net Risk</th>
            <th>Risk Owner</th>
          </tr>
        </thead>
        <tbody id="projectRisks"></tbody>
      </table>
      </div>

      <div class="ProjectIssue">
        <h3>Project Issues</h3>
        <table>
          <thead>
            <tr>
              <th>Issue Ref</th>
              <th>Description</th>
              <th>Escalation</th>
              <th>Action</th>
              <th>Action Date</th>
              <th>Issue Owner</th>
            </tr>
          </thead>
          <tbody id="projectIssues"></tbody>
        </table>
      </div>

      <div class="CompletedThisPeriod">
      <h3>Tasks - completed this period</h3>
      <table>
        <thead class="WithoutBorder">
          <tr>
          <th>Task Description</th>
          <th class="narrow">Due Date</th>
          <th>Assigned to</th>
          </tr>
        </thead>
        <tbody id="completedThisPeriod" class="WithoutBorder"></tbody>
      </table>
    </div>

    <div class="FocusForNextPeriod">
    <h3>Tasks - focus for next period</h3>
      <table>
      <thead>
        <tr>
          <th>Task Description</th>
          <th class="narrow">Due Date</th>
          <th>Assigned to</th>
        </tr>
      </thead>
      <tbody id="focusForNextPeriod"></tbody>
      </table>
    </div>

    <div class="gates">
      <h3>Gates</h3>
      <table>
        <thead>
          <tr>
            <th>Current Gate</th>
            <th>Requested gate</th>
            <th>Gate Comments</th>
            <th>Status</th>
            <th>Date submitted</th>
            <th>Date approved / rejected</th>
          </tr>
        </thead>
        <tbody id="gates">
        </tbody>
      </table>
    </div>

    <div class="lessonsLearned">
      <h3>Lessons Learned</h3>
      <table>
        <thead>
          <tr>
            <th>Situation Description</th>
            <th>Lesson Learned</th>
            <th>Follow up action</th>
            <th class="narrow">Action Date</th>
            <th>Actions by</th>
            <th>Outcomes</th>
          </tr>
        </thead>
        <tbody id="lessonsLearned">
        </tbody>
      </table>
    </div>

    <div class="keyMilestones">
      <h3>Key Milestones</h3>
      <table>
        <thead>
          <tr>
            <th>Milestone Type</th>
            <th>Gate</th>
            <th class="narrow">Date</th>
          </tr>
        </thead>
        <tbody id="keyMilestones">
        </tbody>
      </table>
    </div>

    
    <div class="clear"></div>
    </div>`;

    require('./my-script.js');
    $("nothingSelected").show();
    axios('https://jsonplaceholder.typicode.com/todos/1');
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
