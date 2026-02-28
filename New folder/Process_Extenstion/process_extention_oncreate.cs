using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DUC.ProcessAutomation.Shared.Extensions;
using DUC.ProcessAutomation.Shared.Utilities;
using DUC.ProcessAutomation.Shared.Models.EarlyBound;
using Microsoft.Xrm.Sdk.Messages;
using System.Diagnostics;
using System.Runtime.Remoting.Contexts;
using Microsoft.Crm.Sdk.Messages;

namespace DUC.ProcessAutomation.TargetEntityPlugins
{
	/// <summary>
	/// This Plugin is used to create/update process extension record and set (Regarding & Customer) Only
	/// </summary>
	public class AutoCreateProcessExtension : EntityExtensions, IPlugin
	{
		protected override void Execute()
		{
			if (ExecutionContext.Depth > 1) return;

			duc_ProcessExtension record = null;

			if (IsCreate &&
				(!PrimaryEntity.Contains(duc_ProcessExtension.EntityLogicalName.ToLower())
				|| PrimaryEntity.GetAttributeValue<EntityReference>(duc_ProcessExtension.EntityLogicalName.ToLower()) != null))
			{
				Trace("Handling create event for target entity.");



				record = new duc_ProcessExtension
				{
					RegardingObjectId = PrimaryEntity.ToEntityReference(),

					//duc_processDefinition = processDefinition.ToEntityReference()
				};


			}

			if (record == null)
			{
				Trace("No record to upsert. Exiting.");
				return;
			}

			// Perform the upsert
			Trace(string.Format("Executing upsertRequest"));

			var upsertRequest = new UpsertRequest { Target = record };
			var upsertResponse = (UpsertResponse)OrganizationService.Execute(upsertRequest);
			Trace(string.Format("upsertRequest executded"));



			if (upsertResponse.RecordCreated)
			{
				Trace("duc_ProcessExtension record was created.");
				var entity = new Entity(PrimaryEntity.LogicalName, PrimaryEntity.Id);
				entity[duc_ProcessExtension.EntityLogicalName] = upsertResponse.Target;
				OrganizationService.Update(entity);

			}
			else
				Trace("duc_ProcessExtension record was updated.");
		}

	}
}

