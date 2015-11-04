package com.nbty.plm.px.extbulkchgs;

import com.agile.api.APIException;
import com.agile.api.IAgileSession;
import com.agile.api.INode;
import com.agile.api.IQuery;
import com.agile.api.ITable;
import com.agile.api.ItemConstants;
import com.agile.px.ActionResult;
import com.agile.px.EventActionResult;
import com.agile.px.IEventAction;
import com.agile.px.IEventInfo;

public class ExtractBulkChanges implements IEventAction {

	@Override
	public EventActionResult doAction(IAgileSession session, INode node,
			IEventInfo eventInfo) {
		try {
			
			IQuery query = (IQuery) session.createObject(IQuery.OBJECT_TYPE, ItemConstants.CLASS_PART);
			query.setCaseSensitive(false);
			query.setCriteria("[Item Type] Equal To 'Bulk' ");
			ITable results = query.execute();
			
			return new EventActionResult(eventInfo, new ActionResult(ActionResult.STRING, "Test " + results.isEmpty()));
		} catch (APIException e) {
			return new EventActionResult(eventInfo, new ActionResult(ActionResult.EXCEPTION, e));
		}
	}

}
