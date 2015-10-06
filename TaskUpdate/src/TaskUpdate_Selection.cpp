#include "TaskUpdate_Selection.h"


CTaskUpdate_Selection::CTaskUpdate_Selection(void)
{
}

CTaskUpdate_Selection::~CTaskUpdate_Selection(void)
{

}

void CTaskUpdate_Selection::addProject(TaskUpdate_SelectionProject project)
{
    m_selectionProjectList.push_back(project);
}
void CTaskUpdate_Selection::removeProject(CString GUID)
{
    for (auto itProject = m_selectionProjectList.begin(); itProject != m_selectionProjectList.end(); itProject++)
    {
        if (0 == itProject->GUID.CompareNoCase(GUID))
        {
            m_selectionProjectList.erase(itProject);
            break;
        }
    }
}
bool CTaskUpdate_Selection::projectExists(CString GUID)
{
    bool returnValue = false;
    for (auto itProject = m_selectionProjectList.begin(); itProject != m_selectionProjectList.end(); itProject++)
    {
        if (0 == itProject->GUID.CompareNoCase(GUID))
        {
            returnValue = true;
            break;
        }
    }

    return returnValue;
}
bool CTaskUpdate_Selection::taskIsRemembered(CString GUID, long long taskUID)
{
    bool returnValue = false;
    for (auto itProject = m_selectionProjectList.begin(); itProject != m_selectionProjectList.end(); itProject++)
    {
        if (0 == itProject->GUID.CompareNoCase(GUID))
        {
            for (auto itTask = itProject->taskSelectionList.begin(); itTask != itProject->taskSelectionList.end(); itTask++)
            {
                if (*itTask == taskUID)
                {
                    returnValue = true;
                    break;
                }
            }
            break;
        }
    }

    return returnValue;
}
bool CTaskUpdate_Selection::resourceIsRemembered(CString GUID, long long resourceUID)
{
    bool returnValue = false;
    for (auto itProject = m_selectionProjectList.begin(); itProject != m_selectionProjectList.end(); itProject++)
    {
        if (0 == itProject->GUID.CompareNoCase(GUID))
        {
            for (auto itResource = itProject->resourceSelectionList.begin(); itResource != itProject->resourceSelectionList.end(); itResource++)
            {
                if (*itResource == resourceUID)
                {
                    returnValue = true;
                    break;
                }
            }
            break;
        }
    }

    return returnValue;
}