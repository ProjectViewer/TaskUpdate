#include "TaskUpdate_Utils.h"

void CTU_Utils::GetFilePaths(CString mppFilePath, CString& spvFilePath, CString& updatesDirPath)
{
    CString strUpdatesDirName = _T("Project Updates");

    int pos = mppFilePath.ReverseFind('\\');
    CString fileName = mppFilePath.Mid(pos + 1, mppFilePath.GetLength() - pos - 5);
    updatesDirPath = mppFilePath.Mid(0, pos) + '\\' + strUpdatesDirName;
    spvFilePath = updatesDirPath + '\\' + fileName + _T(".spv");
}
bool CTU_Utils::FileExists(CString filePath)
{
    DWORD dwAttr = GetFileAttributes(filePath);
    if (dwAttr == 0xffffffff) //file does not exists
    {
        DWORD dwError = GetLastError();
        if (dwError == ERROR_FILE_NOT_FOUND)
            return false;
        else //other errors about file existence
            return false;
    }
    else //file exists
        return true;

    return false;
}

bool CTU_Utils::DirExists(CString dirPath)
{
    DWORD dwAttr = GetFileAttributes(dirPath);
    if (dwAttr == 0xffffffff) //dir does not exists
        return false;
    else //dir exists
        return true;
    return false;
}

bool CTU_Utils::LoadXML(CString strPath, CTUProjectReport& theProjectReport, int reportingMode)
{
    CTU_XMLParser theParser;

    VARIANT_BOOL bSuccess = false;
    CComPtr<IXMLDOMDocument> spXMLDOM;
    CString str;

    HRESULT hrCreateDoc = spXMLDOM.CoCreateInstance(__uuidof(DOMDocument));
    if (hrCreateDoc != S_OK || spXMLDOM.p == NULL)
    {
        return false;
    }

    HRESULT hrLoadXML = spXMLDOM->load(CComVariant(strPath), &bSuccess);
    if (hrLoadXML != S_OK || !bSuccess)
    {
        return false;
    }

    CComPtr<IXMLDOMNode> spXMLProjectReportNode;
    HRESULT hrProjectReportNode = spXMLDOM->selectSingleNode(CComBSTR (L"ProjectReport"), &spXMLProjectReportNode);
    if (hrProjectReportNode != S_OK || spXMLProjectReportNode.p == NULL)
    {
        return false;
    }

    if (S_OK != theParser.GetAttributeValue(spXMLProjectReportNode, L"ProjectGUID", str))
    {
        return false;
    }
    theProjectReport.SetProjectGUID(str);

    if (S_OK != theParser.GetNodeValue(spXMLProjectReportNode,  L"ProjectName", str))
    {
        return false;
    }
    theProjectReport.SetProjectName(str);

    if (reportingMode == TUTaskReportingMode)
    {
        CComPtr<IXMLDOMNode> spXMLTaskReportNode;
        HRESULT hrTaskReportNode = spXMLProjectReportNode->selectSingleNode(CComBSTR(L"TaskReport"), &spXMLTaskReportNode);
        while (hrTaskReportNode == S_OK && spXMLTaskReportNode.p != NULL)
        {
            CTUTaskReport* pTaskReport = new CTUTaskReport();

            if (S_OK != theParser.GetAttributeValue(spXMLTaskReportNode, L"TaskUID", str))
            {
                return false;
            }
            pTaskReport->SetTaskUID(_ttoi(str));

            if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"ProjectGUID", str))
            {
                return false;
            }
            pTaskReport->SetProjectGUID(str);

            if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"ProjectName", str))
            {
                return false;
            }
            pTaskReport->SetProjectName(str);

            if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"TaskName", str))
            {
                return false;
            }
            pTaskReport->SetTaskName(str);

            if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"WorkInsertMode", str))
            {
                return false;
            }
            pTaskReport->SetWorkInsertMode((enumTUTaskWorkInsertMode)_ttoi(str));

            if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"ReportGUID", str))
            {
                return false;
            }
            pTaskReport->SetReportGUID(str);

            if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"CreationDate", str))
            {
                return false;
            }
            pTaskReport->SetCreationDate(_ttoi(str));

            if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"CurrentPercentComplete", str))
            {
                return false;
            }
            pTaskReport->SetCurrentPercentComplete(_ttoi(str));

            if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"ReportedPercentComplete", str))
            {
                return false;
            }
            pTaskReport->SetReportedPercentComplete(_ttoi(str));

            if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"CurrentPercentWorkComplete", str))
            {
                return false;
            }
            pTaskReport->SetCurrentPercentWorkComplete(_ttoi(str));

            if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"ReportedPercentWorkComplete", str))
            {
                return false;
            }
            pTaskReport->SetReportedPercentWorkComplete(_ttoi(str));

            if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"CurrentActualWork", str))
            {
                return false;
            }
            pTaskReport->SetCurrentActualWork(_ttoi(str));

            if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"ReportedActualWork", str))
            {
                return false;
            }
            pTaskReport->SetReportedActualWork(_ttoi(str));

			if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"CurrentActualStart", str))
			{
				return false;
			}
			pTaskReport->SetCurrentActualStart(_ttoi(str));

			if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"ReportedActualStart", str))
			{
				return false;
			}
			pTaskReport->SetReportedActualStart(_ttoi(str));

			if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"CurrentActualFinish", str))
			{
				return false;
			}
			pTaskReport->SetCurrentActualFinish(_ttoi(str));

			if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"ReportedActualFinish", str))
			{
				return false;
			}
			pTaskReport->SetReportedActualFinish(_ttoi(str));

            if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"CurrentActualDuration", str))
            {
                return false;
            }
            pTaskReport->SetCurrentActualDuration(_ttoi(str));

            if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"ReportedActualDuration", str))
            {
                return false;
            }
            pTaskReport->SetReportedActualDuration(_ttoi(str));

            if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"CurrentNotes", str))
            {
                return false;
            }
            pTaskReport->SetCurrentNotes(str);

            if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"ReportedNotes", str))
            {
                return false;
            }
            pTaskReport->SetReportedNotes(str);

            if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"CustomText1", str))
            {
                return false;
            }
            pTaskReport->SetCustomText1FieldText(str);

            if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"CustomText2", str))
            {
                return false;
            }
            pTaskReport->SetCustomText2FieldText(str);

            if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"CustomText3", str))
            {
                return false;
            }
            pTaskReport->SetCustomText3FieldText(str);

            if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"TMNote", str))
            {
                return false;
            }
            pTaskReport->SetTMNote(str);

            if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"PMNote", str))
            {
                return false;
            }
            pTaskReport->SetPMNote(str);

            if (S_OK != theParser.GetNodeValue(spXMLTaskReportNode,  L"Status", str))
            {
                return false;
            }
            pTaskReport->SetStatus((enumTUTaskReportsStatus)_ttoi(str));

            theProjectReport.AddTaskReport(pTaskReport);

            CComPtr<IXMLDOMNode> tempTaskNode = spXMLTaskReportNode;
            spXMLTaskReportNode.p = NULL;
            hrTaskReportNode = tempTaskNode->get_nextSibling(&spXMLTaskReportNode);
        }
    }

    if (reportingMode == TUAssignmentReportingMode)
    {
        CComPtr<IXMLDOMNode> spXMLAssignmentReportNode;
        HRESULT hrAssignmentReportNode = spXMLProjectReportNode->selectSingleNode(CComBSTR(L"AssignmentReport"), &spXMLAssignmentReportNode);
        while (hrAssignmentReportNode == S_OK && spXMLAssignmentReportNode.p != NULL)
        {
            CTUAssignmentReport* pAssgnReport = new CTUAssignmentReport();

            if (S_OK != theParser.GetAttributeValue(spXMLAssignmentReportNode, L"AssignmentUID", str))
            {
                return false;
            }
            pAssgnReport->SetAssignmentUID(_ttoi(str));

            if (S_OK != theParser.GetNodeValue(spXMLAssignmentReportNode,  L"ProjectGUID", str))
            {
                return false;
            }
            pAssgnReport->SetProjectGUID(str);

            if (S_OK != theParser.GetNodeValue(spXMLAssignmentReportNode,  L"ProjectName", str))
            {
                return false;
            }
            pAssgnReport->SetProjectName(str);

            if (S_OK != theParser.GetNodeValue(spXMLAssignmentReportNode,  L"TaskUID", str))
            {
                return false;
            }
            pAssgnReport->SetTaskUID(_ttoi(str));

            if (S_OK != theParser.GetNodeValue(spXMLAssignmentReportNode,  L"TaskName", str))
            {
                return false;
            }
            pAssgnReport->SetTaskName(str);

            if (S_OK != theParser.GetNodeValue(spXMLAssignmentReportNode,  L"ResourceUID", str))
            {
                return false;
            }
            pAssgnReport->SetResourceUID(_ttoi(str));

            if (S_OK != theParser.GetNodeValue(spXMLAssignmentReportNode,  L"ResourceName", str))
            {
                return false;
            }
            pAssgnReport->SetResourceName(str);

            if (S_OK != theParser.GetNodeValue(spXMLAssignmentReportNode,  L"WorkInsertMode", str))
            {
                return false;
            }
            pAssgnReport->SetWorkInsertMode((enumTUAssgnWorkInsertMode)_ttoi(str));

            if (S_OK != theParser.GetNodeValue(spXMLAssignmentReportNode,  L"MarkAsCompleted", str))
            {
                return false;
            }
            pAssgnReport->SetMarkedAsCompleted(_ttoi(str) == 1);

            CComPtr<IXMLDOMNode> spXMLDateReportNode;
            HRESULT hrDateReportNode = spXMLAssignmentReportNode->selectSingleNode(CComBSTR(L"DateReport"), &spXMLDateReportNode);
            while (hrDateReportNode == S_OK && spXMLDateReportNode.p != NULL)
            {
                CTUDateReport* pDateReport = new CTUDateReport();

                if (S_OK != theParser.GetAttributeValue(spXMLDateReportNode, L"ReportGUID", str))
                {
                    return false;
                }
                pDateReport->SetReportGUID(str);

                if (S_OK != theParser.GetNodeValue(spXMLDateReportNode,  L"CreationDate", str))
                {
                    return false;
                }
                pDateReport->SetCreationDate(_ttol(str));

                if (S_OK != theParser.GetNodeValue(spXMLDateReportNode,  L"ReportDate", str))
                {
                    return false;
                }
                pDateReport->SetReportDate(_ttol(str));

                if (S_OK != theParser.GetNodeValue(spXMLDateReportNode,  L"PlannedWork", str))
                {
                    return false;
                }
                pDateReport->SetPlannedWork(_ttoi(str));

                if (S_OK != theParser.GetNodeValue(spXMLDateReportNode,  L"ReportedWork", str))
                {
                    return false;
                }
                pDateReport->SetReportedWork(_ttoi(str));

                if (S_OK != theParser.GetNodeValue(spXMLDateReportNode,  L"PlannedOvertimeWork", str))
                {
                    return false;
                }
                pDateReport->SetPlannedOvertimeWork(_ttoi(str));

                if (S_OK != theParser.GetNodeValue(spXMLDateReportNode,  L"ReportedOvertimeWork", str))
                {
                    return false;
                }
                pDateReport->SetReportedOvertimeWork(_ttoi(str));

                if (S_OK != theParser.GetNodeValue(spXMLDateReportNode,  L"CurrentPercentComplete", str))
                {
                    return false;
                }
                pDateReport->SetCurrentPercentComplete(_ttoi(str));

                if (S_OK != theParser.GetNodeValue(spXMLDateReportNode,  L"ReportedPercentComplete", str))
                {
                    return false;
                }
                pDateReport->SetReportedPercentComplete(_ttoi(str));

                if (S_OK != theParser.GetNodeValue(spXMLDateReportNode,  L"TMNote", str))
                {
                    str = _T("");
                }
                pDateReport->SetTMNote(str);

                if (S_OK != theParser.GetNodeValue(spXMLDateReportNode,  L"PMNote", str))
                {
                    str = _T("");
                }
                pDateReport->SetPMNote(str);

                if (S_OK != theParser.GetNodeValue(spXMLDateReportNode,  L"Status", str))
                {
                    return false;
                }
                pDateReport->SetStatus((enumTUReportsStatus)_ttoi(str));

                pAssgnReport->AddDateReport(pDateReport);

                CComPtr<IXMLDOMNode> tempDateNode = spXMLDateReportNode;
                spXMLDateReportNode.p = NULL;
                hrDateReportNode = tempDateNode->get_nextSibling(&spXMLDateReportNode);
            }

            theProjectReport.AddAssgnReport(pAssgnReport);

            CComPtr<IXMLDOMNode> tempAssgnNode = spXMLAssignmentReportNode;
            spXMLAssignmentReportNode.p = NULL;
            hrAssignmentReportNode = tempAssgnNode->get_nextSibling(&spXMLAssignmentReportNode);
        }
    }

    return true;
}
bool CTU_Utils::SaveXML(CString strPath, CTUProjectReport* theProjectReport)
{
    CTU_XMLParser theParser;

    CComVariant varValue(VT_EMPTY);
    VARIANT_BOOL bSuccess = false;
    CComPtr<IXMLDOMDocument> pXMLDom;
    CString str;

    TUDateReportsList::iterator itDateReport;
    TUAssignmentReportsList::iterator itAssgnReport;
    TUTaskReportsList::iterator itTaskReport;

    HRESULT hrCreateDoc = pXMLDom.CoCreateInstance(__uuidof(DOMDocument));
    if (hrCreateDoc != S_OK || pXMLDom.p == NULL)
    {
        return false;
    }

    if (S_OK != theParser.AddProcessingInstructionNode(pXMLDom, L"xml", L"version='1.0' encoding='UTF-16'"))
    {
        return false;
    }

    CComPtr<IXMLDOMNode> spProjectReportNode;
    if (S_OK != theParser.AddRootElementNode(pXMLDom, L"ProjectReport", &spProjectReportNode))
    {
        return false;
    }

    str = theProjectReport->GetProjectGUID();
    if (S_OK != theParser.AddAtributeToNode(pXMLDom, spProjectReportNode, L"ProjectGUID", CComBSTR(str)))
    {
        return false;
    }

    CComPtr<IXMLDOMNode> spProjectNameNode;
    str = theProjectReport->GetProjectName();
    if (S_OK != theParser.AddTextNode(pXMLDom, spProjectReportNode, L"ProjectName", str, &spProjectNameNode))
    {
        return false;
    }

    //task reports
    for (itTaskReport = theProjectReport->GetTaskReportsList()->begin();
        itTaskReport != theProjectReport->GetTaskReportsList()->end();
        itTaskReport++)
    {
        CComPtr<IXMLDOMNode> spTaskReportNode;
        if (S_OK != theParser.AddChildElementNode(pXMLDom, spProjectReportNode, L"TaskReport", &spTaskReportNode))
        {
            return false;
        }

        str.Format(_T("%lld"), (*itTaskReport)->GetTaskUID());
        if (S_OK != theParser.AddAtributeToNode(pXMLDom, spTaskReportNode, L"TaskUID", CComBSTR(str)))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spTaskProjGUIDNode;
        str = (*itTaskReport)->GetProjectGUID();
        if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"ProjectGUID", str, &spTaskProjGUIDNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spTaskProjNameNode;
        str = (*itTaskReport)->GetProjectName();
        if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"ProjectName", str, &spTaskProjNameNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spTaskTaskNameNode;
        str = (*itTaskReport)->GetTaskName();
        if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"TaskName", str, &spTaskTaskNameNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spTaskWorkInsertModeNode;
        str.Format(_T("%d"), (*itTaskReport)->GetWorkInsertMode());
        if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"WorkInsertMode", str, &spTaskWorkInsertModeNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spTaskWorkReportGUIDNode;
        str = (*itTaskReport)->GetReportGUID();
        if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"ReportGUID", str, &spTaskWorkReportGUIDNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spTaskCreationDateNode;
        str.Format(_T("%d"), (*itTaskReport)->GetCreationDate());
        if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"CreationDate", str, &spTaskCreationDateNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spTaskCurrentPercentCompleteNode;
        str.Format(_T("%d"), (*itTaskReport)->GetCurrentPercentComplete());
        if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"CurrentPercentComplete", str, &spTaskCurrentPercentCompleteNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spTaskReportedPercentCompleteNode;
        str.Format(_T("%d"), (*itTaskReport)->GetReportedPercentComplete());
        if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"ReportedPercentComplete", str, &spTaskReportedPercentCompleteNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spTaskCurrentPercentWorkCompleteNode;
        str.Format(_T("%d"), (*itTaskReport)->GetCurrentPercentWorkComplete());
        if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"CurrentPercentWorkComplete", str, &spTaskCurrentPercentWorkCompleteNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spTaskReportedPercentWorkCompleteNode;
        str.Format(_T("%d"), (*itTaskReport)->GetReportedPercentWorkComplete());
        if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"ReportedPercentWorkComplete", str, &spTaskReportedPercentWorkCompleteNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spTaskCurrentActualWorkNode;
        str.Format(_T("%d"), (*itTaskReport)->GetCurrentActualWork());
        if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"CurrentActualWork", str, &spTaskCurrentActualWorkNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spTaskReportedActualWorkNode;
        str.Format(_T("%d"), (*itTaskReport)->GetReportedActualWork());
        if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"ReportedActualWork", str, &spTaskReportedActualWorkNode))
        {
            return false;
        }

		CComPtr<IXMLDOMNode> spTaskCurrentActualStartNode;
		str.Format(_T("%d"), (*itTaskReport)->GetCurrentActualStart());
		if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"CurrentActualStart", str, &spTaskCurrentActualStartNode))
		{
			return false;
		}

		CComPtr<IXMLDOMNode> spTaskReportedActualStartNode;
		str.Format(_T("%d"), (*itTaskReport)->GetReportedActualStart());
		if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"ReportedActualStart", str, &spTaskReportedActualStartNode))
		{
			return false;
		}

		CComPtr<IXMLDOMNode> spTaskCurrentActualFinishNode;
		str.Format(_T("%d"), (*itTaskReport)->GetCurrentActualFinish());
		if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"CurrentActualFinish", str, &spTaskCurrentActualFinishNode))
		{
			return false;
		}

		CComPtr<IXMLDOMNode> spTaskReportedActualFinishNode;
		str.Format(_T("%d"), (*itTaskReport)->GetReportedActualFinish());
		if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"ReportedActualFinish", str, &spTaskReportedActualFinishNode))
		{
			return false;
		}

        CComPtr<IXMLDOMNode> spTaskCurrentActualDurationNode;
        str.Format(_T("%d"), (*itTaskReport)->GetCurrentActualDuration());
        if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"CurrentActualDuration", str, &spTaskCurrentActualDurationNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spTaskReportedActualDurationNode;
        str.Format(_T("%d"), (*itTaskReport)->GetReportedActualDuration());
        if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"ReportedActualDuration", str, &spTaskReportedActualDurationNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spTaskCurrentNotesNode;
        str = (*itTaskReport)->GetCurrentNotes();
        if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"CurrentNotes", str, &spTaskCurrentNotesNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spTaskReportedNotesNode;
        str = (*itTaskReport)->GetReportedNotes();
        if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"ReportedNotes", str, &spTaskReportedNotesNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spTaskReportedCustomText1Node;
        str = (*itTaskReport)->GetCustomText1FieldText();
        if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"CustomText1", str, &spTaskReportedCustomText1Node))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spTaskReportedCustomText2Node;
        str = (*itTaskReport)->GetCustomText2FieldText();
        if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"CustomText2", str, &spTaskReportedCustomText2Node))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spTaskReportedCustomText3Node;
        str = (*itTaskReport)->GetCustomText3FieldText();
        if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"CustomText3", str, &spTaskReportedCustomText3Node))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spTaskTMNoteNode;
        str = (*itTaskReport)->GetTMNote();
        if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"TMNote", str, &spTaskTMNoteNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spTaskPMNoteNode;
        str = (*itTaskReport)->GetPMNote();
        if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"PMNote", str, &spTaskPMNoteNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spTaskStatusNode;
        str.Format(_T("%d"), (*itTaskReport)->GetStatus());
        if (S_OK != theParser.AddTextNode(pXMLDom, spTaskReportNode, L"Status", str, &spTaskStatusNode))
        {
            return false;
        }
    }

	//assignment reports
    for (itAssgnReport = theProjectReport->GetAssignmentReportsList()->begin();
         itAssgnReport != theProjectReport->GetAssignmentReportsList()->end();
         itAssgnReport++)
    {
        CComPtr<IXMLDOMNode> spAssignemntReportNode;
        if (S_OK != theParser.AddChildElementNode(pXMLDom, spProjectReportNode, L"AssignmentReport", &spAssignemntReportNode))
        {
            return false;
        }

        str.Format(_T("%lld"), (*itAssgnReport)->GetAssignmentUID());
        if (S_OK != theParser.AddAtributeToNode(pXMLDom, spAssignemntReportNode, L"AssignmentUID", CComBSTR(str)))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spAssgnProjGUIDNode;
        str = (*itAssgnReport)->GetProjectGUID();
        if (S_OK != theParser.AddTextNode(pXMLDom, spAssignemntReportNode, L"ProjectGUID", str, &spAssgnProjGUIDNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spAssgnProjNameNode;
        str = (*itAssgnReport)->GetProjectName();
        if (S_OK != theParser.AddTextNode(pXMLDom, spAssignemntReportNode, L"ProjectName", str, &spAssgnProjNameNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spAssgnResNameNode;
        str = (*itAssgnReport)->GetResourceName();
        if (S_OK != theParser.AddTextNode(pXMLDom, spAssignemntReportNode, L"ResourceName", str, &spAssgnResNameNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spAssgnResUIDNode;
        str.Format(_T("%lld"), (*itAssgnReport)->GetResourceUID());
        if (S_OK != theParser.AddTextNode(pXMLDom, spAssignemntReportNode, L"ResourceUID", str, &spAssgnResUIDNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spAssgnTaskNameNode;
        str = (*itAssgnReport)->GetTaskName();
        if (S_OK != theParser.AddTextNode(pXMLDom, spAssignemntReportNode, L"TaskName", str, &spAssgnTaskNameNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spAssgnTaskUIDNode;
        str.Format(_T("%lld"), (*itAssgnReport)->GetTaskUID());
        if (S_OK != theParser.AddTextNode(pXMLDom, spAssignemntReportNode, L"TaskUID", str, &spAssgnTaskUIDNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spAssgnMarkAsComplNode;
        str.Format(_T("%d"), (*itAssgnReport)->IsMarkedAsCompleted());
        if (S_OK != theParser.AddTextNode(pXMLDom, spAssignemntReportNode, L"MarkAsCompleted", str, &spAssgnMarkAsComplNode))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spAssgnWorkInsertModeNode;
        str.Format(_T("%d"), (*itAssgnReport)->GetWorkInsertMode());
        if (S_OK != theParser.AddTextNode(pXMLDom, spAssignemntReportNode, L"WorkInsertMode", str, &spAssgnWorkInsertModeNode))
        {
            return false;
        }

        for (itDateReport = (*itAssgnReport)->GetDateReportsList()->begin();
             itDateReport != (*itAssgnReport)->GetDateReportsList()->end();
             itDateReport++)
        {
            CComPtr<IXMLDOMNode> spDateReportNode;
            if (S_OK != theParser.AddChildElementNode(pXMLDom, spAssignemntReportNode, L"DateReport", &spDateReportNode))
            {
                return false;
            }

            str = (*itDateReport)->GetReportGUID();
            if (S_OK != theParser.AddAtributeToNode(pXMLDom, spDateReportNode, L"ReportGUID", CComBSTR(str)))
            {
                return false;
            }

            CComPtr<IXMLDOMNode> spDateCreationDateNode;
            str.Format(_T("%d"), (*itDateReport)->GetCreationDate());
            if (S_OK != theParser.AddTextNode(pXMLDom, spDateReportNode, L"CreationDate", str, &spDateCreationDateNode))
            {
                return false;
            }

            CComPtr<IXMLDOMNode> spDateReportDateNode;
            str.Format(_T("%d"), (*itDateReport)->GetReportDate());
            if (S_OK != theParser.AddTextNode(pXMLDom, spDateReportNode, L"ReportDate", str, &spDateReportDateNode))
            {
                return false;
            }

            CComPtr<IXMLDOMNode> spDatePlannedWorkNode;
            str.Format(_T("%d"), (*itDateReport)->GetPlannedWork());
            if (S_OK != theParser.AddTextNode(pXMLDom, spDateReportNode, L"PlannedWork", str, &spDatePlannedWorkNode))
            {
                return false;
            }

            CComPtr<IXMLDOMNode> spDateReportedWorkNode;
            str.Format(_T("%d"), (*itDateReport)->GetReportedWork());
            if (S_OK != theParser.AddTextNode(pXMLDom, spDateReportNode, L"ReportedWork", str, &spDateReportedWorkNode))
            {
                return false;
            }

            CComPtr<IXMLDOMNode> spDatePlannedOvertimeWorkNode;
            str.Format(_T("%d"), (*itDateReport)->GetPlannedOvertimeWork());
            if (S_OK != theParser.AddTextNode(pXMLDom, spDateReportNode, L"PlannedOvertimeWork", str, &spDatePlannedOvertimeWorkNode))
            {
                return false;
            }

            CComPtr<IXMLDOMNode> spDateReportedOvertimeWorkNode;
            str.Format(_T("%d"), (*itDateReport)->GetReportedOvertimeWork());
            if (S_OK != theParser.AddTextNode(pXMLDom, spDateReportNode, L"ReportedOvertimeWork", str, &spDateReportedOvertimeWorkNode))
            {
                return false;
            }

            CComPtr<IXMLDOMNode> spDateCurrentPercComplNode;
            str.Format(_T("%d"), (*itDateReport)->GetCurrentPercentComplete());
            if (S_OK != theParser.AddTextNode(pXMLDom, spDateReportNode, L"CurrentPercentComplete", str, &spDateCurrentPercComplNode))
            {
                return false;
            }

            CComPtr<IXMLDOMNode> spDateReportedPercComplNode;
            str.Format(_T("%d"), (*itDateReport)->GetReportedPercentComplete());
            if (S_OK != theParser.AddTextNode(pXMLDom, spDateReportNode, L"ReportedPercentComplete", str, &spDateReportedPercComplNode))
            {
                return false;
            }

            CComPtr<IXMLDOMNode> spDateTMNoteNode;
            str = (*itDateReport)->GetTMNote();
            if (S_OK != theParser.AddTextNode(pXMLDom, spDateReportNode, L"TMNote", str, &spDateTMNoteNode))
            {
                return false;
            }

            CComPtr<IXMLDOMNode> spDatePMNoteNode;
            str = (*itDateReport)->GetPMNote();
            if (S_OK != theParser.AddTextNode(pXMLDom, spDateReportNode, L"PMNote", str, &spDatePMNoteNode))
            {
                return false;
            }

            CComPtr<IXMLDOMNode> spDateStatusNode;
            str.Format(_T("%d"), (*itDateReport)->GetStatus());
            if (S_OK != theParser.AddTextNode(pXMLDom, spDateReportNode, L"Status", str, &spDateStatusNode))
            {
                return false;
            }
        }
    }

    if (S_OK != pXMLDom->save(CComVariant(strPath)))
    {
        return false;
    }

    return true;
}

bool CTU_Utils::IsValidGUID(CString strGUID)
{
    bool isValidGUID = true;
    int i = 0;
    if (strGUID.GetLength() == 36)
    {
        while (i < strGUID.GetLength())
        {
            if (i == 8 || i == 13 || i == 18 || i == 23)
            {
                if (strGUID.GetAt(i) != '-')
                {
                    return false;
                }
            }
            else
            {
                if ((strGUID.GetAt(i) < '0' && strGUID.GetAt(i) > '9') ||
                    (strGUID.GetAt(i) < 'A' && strGUID.GetAt(i) > 'F'))

                {
                    return false;
                }
            }
            i++;
        }
    }
    else
    {
        return false;
    }

    return true;
}

bool CTU_Utils::IsNumber(CString s)
{
    if (s == _T("0") || s == _T("1") || s == _T("2") || s == _T("3") || s == _T("4") || s == _T("5") || s == _T("6") || s == _T("7") || s == _T("8") || s == _T("9"))
        return true;
    return false;
}

CString CTU_Utils::GenerateGUID()
{
    GUID guidID;
    CString strGUID = _T("");
    if (S_OK == CoCreateGuid(&guidID))
    {
        strGUID.Format(_T("%08X-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X"),
                       guidID.Data1,
                       guidID.Data2,
                       guidID.Data3,
                       guidID.Data4[0],
                       guidID.Data4[1],
                       guidID.Data4[2],
                       guidID.Data4[3],
                       guidID.Data4[4],
                       guidID.Data4[5],
                       guidID.Data4[6],
                       guidID.Data4[7]);
    }
    return strGUID.MakeUpper();
}

bool CTU_Utils::IsFileTUActive(CString path)
{
    CString strGUID =  GetCollIDFromFile(path);
    return (CTU_Utils::IsValidGUID(strGUID) && GetIsActiveFromFile(path));
}
CString CTU_Utils::GetCollIDFromFile(CString path)
{
    CDocCustomProperties docprop;
    docprop.LoadCustomProperties(path);
    return docprop.GetCustomProperertyValueFor(_T("Collaboration ID")).MakeUpper();
}
int CTU_Utils::GetReportingModeFromFile(CString path)
{
    CDocCustomProperties docprop;
    docprop.LoadCustomProperties(path);
    CString strMode = docprop.GetCustomProperertyValueFor(_T("Reporting Mode"));
    if (strMode.IsEmpty())
        strMode = _T("1");//default to assignment updating, for old files
    return _ttoi(strMode);
}
bool CTU_Utils::GetIsActiveFromFile(CString path)
{
    CDocCustomProperties docprop;
    docprop.LoadCustomProperties(path);
    CString strActive = docprop.GetCustomProperertyValueFor(_T("IsTUActive"));
    return strActive == _T("1");
}
CString CTU_Utils::GetCustomPropertyValue(CString path, CString name)
{
    CDocCustomProperties docprop;
    docprop.LoadCustomProperties(path);
    return docprop.GetCustomProperertyValueFor(name);
}
enumTUUnit CTU_Utils::GetUnit(TUUnitsMap* theUnitsList, CString str)
{
    TUUnitsMap::const_iterator itUnit;
    std::vector<CString>::const_iterator itUnitVar;
    for (itUnit = theUnitsList->begin(); itUnit != theUnitsList->end(); ++itUnit)
    {
        for (itUnitVar = itUnit->second.begin(); itUnitVar != itUnit->second.end(); ++itUnitVar)
        {
            if (*itUnitVar == str)
                return itUnit->first;
        }
    }

    return TUUnitNone;
}

CString CTU_Utils::FormatFloat(double val)
{
    CString string;
    wchar_t temp[80];

    GetLocaleInfo(NULL, LOCALE_SDECIMAL, temp, 4);
    CString decimalSeparator = temp;

    string.Format(_T("%.2f"), val);
    string.Replace(_T("."), decimalSeparator);
    for (int i = string.GetLength() - 1; i >= 0; i--)
    {
        if (string.GetAt(i) == '0')
            string.SetAt(i, '*');
        else if (string.GetAt(i) == decimalSeparator)
        {
            string.SetAt(i, '*');
            break;
        }
        else
            break;
    }
    string.Remove('*');

    if (string == _T(""))
        string = _T("0");

    return string;
}

bool CTU_Utils::IsFileWriteble(CString path)
{
    bool swriteble = false;
    CFile cfile_object;
    if (cfile_object.Open(path, CFile::modeWrite))
    {
        swriteble = true;
        cfile_object.Close();
    }
    return swriteble;
}

bool CTU_Utils::MPPFilesExist(TUProjectsInfoList* pProjInfoList)
{
    bool bFileExists = true;
    TUProjectsInfoList::iterator itPrjInfo;
    for (itPrjInfo = pProjInfoList->begin(); itPrjInfo != pProjInfoList->end(); ++itPrjInfo)
    {
        if ((*itPrjInfo)->HasChilds())
        {
            bFileExists = bFileExists && MPPFilesExist((*itPrjInfo)->GetSubProjectsList());
        }

        bFileExists = bFileExists && (*itPrjInfo)->MPPFileExists();
    }

    return bFileExists;
}
bool CTU_Utils::SPVFilesExist(TUProjectsInfoList* pProjInfoList)
{
    bool bFileExists = true;
    TUProjectsInfoList::iterator itPrjInfo;
    for (itPrjInfo = pProjInfoList->begin(); itPrjInfo != pProjInfoList->end(); ++itPrjInfo)
    {
        if ((*itPrjInfo)->HasChilds())
        {
            bFileExists = bFileExists && SPVFilesExist((*itPrjInfo)->GetSubProjectsList());
        }

        bFileExists = bFileExists && (*itPrjInfo)->SPVFileExists();
    }

    return bFileExists;
}
bool CTU_Utils::AreSPVFilesWritable(TUProjectsInfoList* pProjInfoList)
{
    bool bFilesWritable = true;
    TUProjectsInfoList::iterator itPrjInfo;
    for (itPrjInfo = pProjInfoList->begin(); itPrjInfo != pProjInfoList->end(); ++itPrjInfo)
    {
        if ((*itPrjInfo)->HasChilds())
        {
            bFilesWritable = bFilesWritable && AreSPVFilesWritable((*itPrjInfo)->GetSubProjectsList());
        }

        bFilesWritable = bFilesWritable && IsFileWriteble((*itPrjInfo)->GetProjectSPVPath());
    }

    return bFilesWritable;
}

bool CTU_Utils::DeleteSPVFiles(TUProjectsInfoList* pProjInfoList)
{
    bool bFileDeleted = true;
    TUProjectsInfoList::iterator itPrjInfo;
    for (itPrjInfo = pProjInfoList->begin(); itPrjInfo != pProjInfoList->end(); ++itPrjInfo)
    {
        if ((*itPrjInfo)->HasChilds())
        {
            bFileDeleted = bFileDeleted && DeleteSPVFiles((*itPrjInfo)->GetSubProjectsList());
        }

        bFileDeleted = bFileDeleted && (*itPrjInfo)->DeleteSPVFile();
    }
    return bFileDeleted;
}
unsigned int CTU_Utils::GetCustomText1FieldID(CString path)
{
    CDocCustomProperties docprop;
    docprop.LoadCustomProperties(path);
    CString strMode = docprop.GetCustomProperertyValueFor(_T("CustomText1"));
    if (strMode.IsEmpty())
    {
        strMode = _T("0");
    }
    return _ttoi(strMode);
}
unsigned int CTU_Utils::GetCustomText2FieldID(CString path)
{
    CDocCustomProperties docprop;
    docprop.LoadCustomProperties(path);
    CString strMode = docprop.GetCustomProperertyValueFor(_T("CustomText2"));
    if (strMode.IsEmpty())
    {
        strMode = _T("0");
    }
    return _ttoi(strMode);
}
unsigned int CTU_Utils::GetCustomText3FieldID(CString path)
{
    CDocCustomProperties docprop;
    docprop.LoadCustomProperties(path);
    CString strMode = docprop.GetCustomProperertyValueFor(_T("CustomText3"));
    if (strMode.IsEmpty())
    {
        strMode = _T("0");
    }
    return _ttoi(strMode);
}

bool CTU_Utils::SaveSelectionXML(CString path, CTaskUpdate_Selection selection)
{
    CTU_XMLParser theParser;

    CComVariant varValue(VT_EMPTY);
    VARIANT_BOOL bSuccess = false;
    CComPtr<IXMLDOMDocument> pXMLDom;
    CString str;

    HRESULT hrCreateDoc = pXMLDom.CoCreateInstance(__uuidof(DOMDocument));
    if (hrCreateDoc != S_OK || pXMLDom.p == NULL)
    {
        return false;
    }

    if (S_OK != theParser.AddProcessingInstructionNode(pXMLDom, L"xml", L"version='1.0' encoding='UTF-16'"))
    {
        return false;
    }

    CComPtr<IXMLDOMNode> spProjectsNode;
    if (S_OK != theParser.AddRootElementNode(pXMLDom, L"Projects", &spProjectsNode))
    {
        return false;
    
    }

    for (auto itSelection = selection.m_selectionProjectList.begin(); itSelection != selection.m_selectionProjectList.end(); itSelection++)
    {
        CComPtr<IXMLDOMNode> spProjectNode;
        if (S_OK != theParser.AddChildElementNode(pXMLDom, spProjectsNode, L"Project", &spProjectNode))
        {
            return false;
        }

        if (S_OK != theParser.AddAtributeToNode(pXMLDom, spProjectNode, L"ProjectGUID", CComBSTR(itSelection->GUID)))
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spTasksListNode;
        if (S_OK != theParser.AddChildElementNode(pXMLDom, spProjectNode, L"Tasks", &spTasksListNode))
        {
            return false;
        }

        for (auto itTask = itSelection->taskSelectionList.begin() ; itTask != itSelection->taskSelectionList.end(); itTask++)
        {
            CComPtr<IXMLDOMNode> spTaskNode;
            str.Format(_T("%lld"), *itTask);
            if (S_OK != theParser.AddChildElementNode(pXMLDom, spTasksListNode, L"Task", &spTaskNode))
            {
                return false;
            }

            if (S_OK != theParser.AddAtributeToNode(pXMLDom, spTaskNode, L"UID", CComBSTR(str)))
            {
                return false;
            }
        }

        CComPtr<IXMLDOMNode> spResourceListNode;
        if (S_OK != theParser.AddChildElementNode(pXMLDom, spProjectNode, L"Resources", &spResourceListNode))
        {
            return false;
        }

        for (auto itResource = itSelection->resourceSelectionList.begin(); itResource != itSelection->resourceSelectionList.end(); itResource++)
        {
            CComPtr<IXMLDOMNode> spResourceNode;
            str.Format(_T("%lld"), *itResource);
            if (S_OK != theParser.AddChildElementNode(pXMLDom, spResourceListNode, L"Resource", &spResourceNode))
            {
                return false;
            }

            if (S_OK != theParser.AddAtributeToNode(pXMLDom, spResourceNode, L"UID", CComBSTR(str)))
            {
                return false;
            }
        }
    }


    if (S_OK != pXMLDom->save(CComVariant(path)))
    {
        return false;
    }

    return true;
}
bool CTU_Utils::LoadSelectionXML(CString path, CTaskUpdate_Selection& selection)
{
    CTU_XMLParser theParser;

    VARIANT_BOOL bSuccess = false;
    CComPtr<IXMLDOMDocument> spXMLDOM;
    CString str;

    HRESULT hrCreateDoc = spXMLDOM.CoCreateInstance(__uuidof(DOMDocument));
    if (hrCreateDoc != S_OK || spXMLDOM.p == NULL)
    {
        return false;
    }

    HRESULT hrLoadXML = spXMLDOM->load(CComVariant(path), &bSuccess);
    if (hrLoadXML != S_OK || !bSuccess)
    {
        return false;
    }

    CComPtr<IXMLDOMNode> spXMLProjectsNode;
    HRESULT hrProjectsNode = spXMLDOM->selectSingleNode(CComBSTR (L"Projects"), &spXMLProjectsNode);
    if (hrProjectsNode != S_OK || spXMLProjectsNode.p == NULL)
    {
        return false;
    }

    CComPtr<IXMLDOMNode> spXMLProjectNode;
    HRESULT hrProjectNode = spXMLProjectsNode->selectSingleNode(CComBSTR(L"Project"), &spXMLProjectNode);
    while (hrProjectNode == S_OK && spXMLProjectNode.p != NULL)
    {
        CTaskUpdate_Selection::TaskUpdate_SelectionProject selectionProject;

        if (S_OK != theParser.GetAttributeValue(spXMLProjectNode, L"ProjectGUID", str))
        {
            return false;
        }
        selectionProject.GUID = str;

        CComPtr<IXMLDOMNode> spXMLTasksNode;
        HRESULT hrTasksNode = spXMLProjectNode->selectSingleNode(CComBSTR(L"Tasks"), &spXMLTasksNode);
        if (hrTasksNode != S_OK || spXMLTasksNode.p == NULL)
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spXMLTaskNode;
        HRESULT hrTaskNode = spXMLTasksNode->selectSingleNode(CComBSTR(L"Task"), &spXMLTaskNode);
        while (hrTaskNode == S_OK && spXMLTaskNode.p != NULL)
        {
            if (S_OK != theParser.GetAttributeValue(spXMLTaskNode, L"UID", str))
            {
                return false;
            }
            selectionProject.taskSelectionList.push_back(_ttoi(str));

            CComPtr<IXMLDOMNode> tempTaskNode = spXMLTaskNode;
            spXMLTaskNode.p = NULL;
            hrTaskNode = tempTaskNode->get_nextSibling(&spXMLTaskNode);
        }

        CComPtr<IXMLDOMNode> spXMLResourcesNode;
        HRESULT hrResourcesNode = spXMLProjectNode->selectSingleNode(CComBSTR(L"Resources"), &spXMLResourcesNode);
        if (hrResourcesNode != S_OK || spXMLResourcesNode.p == NULL)
        {
            return false;
        }

        CComPtr<IXMLDOMNode> spXMLResourceNode;
        HRESULT hrResourceNode = spXMLResourcesNode->selectSingleNode(CComBSTR(L"Resource"), &spXMLResourceNode);
        while (hrResourceNode == S_OK && spXMLResourceNode.p != NULL)
        {
            if (S_OK != theParser.GetAttributeValue(spXMLResourceNode, L"UID", str))
            {
                return false;
            }
            selectionProject.resourceSelectionList.push_back(_ttoi(str));

            CComPtr<IXMLDOMNode> tempResourceNode = spXMLResourceNode;
            spXMLResourceNode.p = NULL;
            hrResourceNode = tempResourceNode->get_nextSibling(&spXMLResourceNode);
        }

        selection.addProject(selectionProject);

        CComPtr<IXMLDOMNode> tempProjectNode = spXMLProjectNode;
        spXMLProjectNode.p = NULL;
        hrProjectNode = tempProjectNode->get_nextSibling(&spXMLProjectNode);
    }

    return true;
}