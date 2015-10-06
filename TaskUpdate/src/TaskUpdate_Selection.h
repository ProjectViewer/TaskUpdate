#pragma once
#include <vector>

class CTaskUpdate_Selection
{
public:
    CTaskUpdate_Selection(void);
    ~CTaskUpdate_Selection(void);

public:

    struct TaskUpdate_SelectionProject
    {
        CString GUID;
        std::vector<long long> taskSelectionList;
        std::vector<long long> resourceSelectionList;
    };

    std::vector<TaskUpdate_SelectionProject> m_selectionProjectList;

    void addProject(TaskUpdate_SelectionProject project);
    void removeProject(CString GUID);
    bool projectExists(CString GUID);
    bool taskIsRemembered(CString GUID, long long taskUID);
    bool resourceIsRemembered(CString GUID, long long resourceUID);
};
