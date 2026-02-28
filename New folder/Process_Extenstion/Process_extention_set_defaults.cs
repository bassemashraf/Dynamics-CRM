using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DUC.ProcessAutomation.Shared.Extensions;
using DUC.ProcessAutomation.Shared.Models.EarlyBound;
using System.Windows.Shapes;
using DUC.ProcessAutomation.Shared.Models.Domain;
using System.Xml;
using Microsoft.Xrm.Sdk.Workflow;

namespace DUC.ProcessAutomation.Plugins.ProcessExtension.ProcessManagement
{
	public class Create_PostOperation_SetDefaults : EntityExtensions<duc_ProcessExtension>, IPlugin
	{
		protected override void Execute()
		{
			if (PrimaryEntity.RegardingObjectId == null)
			{
				Trace("No matching regarding found. Skipping execution.");
				return;
			}

			// Retrieve the matching process definition
			var processDefinition = Utils.GetProcessDefinitionForEntity(PrimaryEntity);

			if (processDefinition == null)
			{
				Trace("No matching process definition found. Skipping execution.");
				return;
			}

			Trace("We got definition");
			var targetRef = PrimaryEntity.GetAttributeValue<EntityReference>("regardingobjectid");
			Trace("we got regarding");

			var columns = new List<string>();

			if (!string.IsNullOrEmpty(processDefinition.duc_TargetEntityCustomerLookupName))
				columns.Add(processDefinition.duc_TargetEntityCustomerLookupName);

			if (!string.IsNullOrEmpty(processDefinition.duc_TargetEntitySubjectName))
				columns.Add(processDefinition.duc_TargetEntitySubjectName);

			if (!string.IsNullOrEmpty(processDefinition.duc_ParentLookup))
				columns.Add(processDefinition.duc_ParentLookup);

			Trace("trying to retrieve");

			var target = OrganizationService.Retrieve(
				targetRef.LogicalName,
				targetRef.Id,
				new ColumnSet(columns.ToArray())
			);
			Trace("after retrieve");
			if (!string.IsNullOrEmpty(processDefinition.duc_TargetEntityCustomerLookupName))
			{
				PrimaryEntity.duc_CustomerId = target.GetAttributeValue<EntityReference>(processDefinition.duc_TargetEntityCustomerLookupName);
			}
			if (!string.IsNullOrEmpty(processDefinition.duc_TargetEntitySubjectName))
			{
				PrimaryEntity.Subject = target.GetAttributeValue<string>(processDefinition.duc_TargetEntitySubjectName);
			}
			var ParentPE = PrimaryEntity.ToEntityReference();

			if (!string.IsNullOrEmpty(processDefinition.duc_ParentLookup))
			{
				var parentTarget = target.GetAttributeValue<EntityReference>(processDefinition.duc_ParentLookup);
				if (parentTarget != null)
				{

					var query = new QueryExpression(duc_ProcessExtension.EntityLogicalName)
					{
						ColumnSet = new ColumnSet(duc_ProcessExtension.PrimaryIdAttribute, nameof(duc_ProcessExtension.duc_processDefinition).ToLower()),
						Criteria = new FilterExpression(LogicalOperator.And)
						{
							Conditions =
						{
							new ConditionExpression("regardingobjectid", ConditionOperator.Equal, parentTarget.Id)
						}
						},
						TopCount = 1
					};

					Trace("Getting the process ext entity");
					var parentPEentity = OrganizationService.RetrieveMultiple(query).Entities.FirstOrDefault();
					if (parentPEentity != null)
					{
						ParentPE = parentPEentity.ToEntityReference();
					}

				}
			}

			PrimaryEntity.duc_ParentProcessExtention = ParentPE;

			//If Create, Set Default Values, Working as preoperation
			if (IsCreate)
			{
				Trace("Handling create event for process extension entity.");

				Trace("Set Default Values");

				//Set Default Values
				PrimaryEntity.duc_processDefinition = processDefinition.ToEntityReference();
				PrimaryEntity.duc_CurrentStage = processDefinition.duc_StartStage;
				//if (processDefinition.duc_StartAction != null)
				//{
				//	PrimaryEntity.duc_LastActionTaken = processDefinition.duc_StartAction;
				//}
				PrimaryEntity.duc_Status = processDefinition.duc_DefaultStatus;
				PrimaryEntity.duc_SubStatus = processDefinition.duc_DefaultSubStatus;
				PrimaryEntity.duc_AssignmentDate = DateTime.UtcNow;
			}
			OrganizationService.Update(PrimaryEntity);


			if (processDefinition.duc_StartAction != null)
			{
				PrimaryEntity.duc_LastActionTaken = processDefinition.duc_StartAction;
				Utils.ProcessUtils.HandleActionChange(PrimaryEntity);
			}

		}
	}
}
