#pragma once
#include <vector>

enum enumTUTaskWorkInsertMode
{
    TUTaskUndefiniedReport          = 0,
    TUTaskPercentCompleteReport     = 1,
    TUTaskPercentWorkCompleteReport = 2,
    TUTaskActualWorkReport          = 3,
	TUTaskActualStartReport			= 4,
	TUTaskActualFinishReport		= 5,
    TUTaskActualDurationReport      = 6,
};

enum enumTUTaskReportsStatus
{
    TUTaskOpen                  = 0,
    TUTaskRejected              = 1,
    TUTaskAccepted              = 2,
    TUTaskPending               = 3,
    TUTaskModified              = 4,
    TUTaskToBeAccepted          = 5,
    TUTaskToBeRejected          = 6,
    TUTaskAll                   = 7
};

class CTUTaskReport
{
public:
    CTUTaskReport(void);
    ~CTUTaskReport(void);

private:
    CString m_reportGUID;
    CTime m_CreationDate;

    int m_CurrentPercentComplete;
    int m_ReportedPercentComplete;
    int m_CurrentPercentWorkComplete;
    int m_ReportedPercentWorkComplete;
    int m_CurrentActualWork;
    int m_ReportedActualWork;
	CTime m_CurrentActualStart;
	CTime m_ReportedActualStart;
	CTime m_CurrentActualFinish;
	CTime m_ReportedActualFinish;
    int m_CurrentActualDuration;
    int m_ReportedActualDuration;

   CString m_CurrentNotes;
   CString m_ReportedNotes;

    CString m_CustomText1FieldText;
    CString m_CustomText2FieldText;
    CString m_CustomText3FieldText;

    CString m_TMNote;
    CString m_PMNote;

    std::vector<CRect> m_PercRect;
    std::vector<CRect> m_PercWorkRect;
    std::vector<CRect> m_ActWorkRect;
	std::vector<CRect> m_ActStartRect;
	std::vector<CRect> m_ActFinishRect;
    std::vector<CRect> m_ActDurationRect;
    std::vector<CRect> m_NotesRect;
    std::vector<CRect> m_CustomText1Rect;
    std::vector<CRect> m_CustomText2Rect;
    std::vector<CRect> m_CustomText3Rect;

    enumTUTaskReportsStatus m_Status;
    bool m_PercentCompleteEdited;             
    bool m_PercentWorkCompleteEdited;         
    bool m_ActualWorkEdited;                  
	bool m_ActualStartEdited;                 
	bool m_ActualFinishEdited;                
    bool m_ActualDurationEdited;        
    bool m_NotesEdited;   
    bool m_CustomText1Edited;                 
    bool m_CustomText2Edited;       
    bool m_CustomText3Edited;       
    bool m_Active;                  
    bool m_IsSummary;

    bool m_Saved;                   
    bool m_CustomTextSaved;
    bool m_NotesSaved;
    bool m_Disabled;                
    UINT_PTR m_UINT_PTR;               //to store pointer address for tooltips
    bool m_Selected;                  


    long long m_TaskUID;
    CString m_ProjectName;
    CString m_ProjectGUID;
    long long m_TaskBaseUID;
    CString m_TaskName;
    enumTUTaskWorkInsertMode m_WorkInsertMode;

    unsigned int m_CustomText1FieldID;
    unsigned int m_CustomText2FieldID;
    unsigned int m_CustomText3FieldID;

public:
    CString GetReportGUID() const;
    void SetReportGUID(CString strReportGUID);
    CTime GetCreationDate() const;
    void SetCreationDate(CTime tCreationDate);

    int GetCurrentPercentComplete() const;
    void SetCurrentPercentComplete(int iCurrentPercentComplete);
    int GetReportedPercentComplete() const;
    void SetReportedPercentComplete(int iReportedPercentComplete);
    int GetCurrentPercentWorkComplete() const;
    void SetCurrentPercentWorkComplete(int iCurrentPercentWorkComplete);
    int GetReportedPercentWorkComplete() const;
    void SetReportedPercentWorkComplete(int iReportedPercentWorkComplete);
    int GetCurrentActualWork() const;
    void SetCurrentActualWork(int iCurrentActualWork);
    int GetReportedActualWork() const;
	void SetReportedActualWork(int iReportedActualWork);
	CTime GetCurrentActualStart() const;
	void SetCurrentActualStart(CTime iCurrentActualStart);
	CTime GetReportedActualStart() const;
	void SetReportedActualStart(CTime iReportedActualStart);
	CTime GetCurrentActualFinish() const;
	void SetCurrentActualFinish(CTime iCurrentActualFinish);
	CTime GetReportedActualFinish() const;
	void SetReportedActualFinish(CTime iReportedActualFinish);
    int GetCurrentActualDuration() const;
    void SetCurrentActualDuration(int iCurrentActualWork);
    int GetReportedActualDuration() const;
    void SetReportedActualDuration(int iReportedActualWork);
    CString GetCurrentNotes() const;
    void SetCurrentNotes(CString currentNotes);
    CString GetReportedNotes() const;
    void SetReportedNotes(CString reportedNotes);

    CString GetCustomText1FieldText() const;
    void SetCustomText1FieldText(CString text);
    CString GetCustomText2FieldText() const;
    void SetCustomText2FieldText(CString text);
    CString GetCustomText3FieldText() const;
    void SetCustomText3FieldText(CString text);
    CString GetCustomTextFieldText(unsigned int fieldID) const;
    void SetCustomTextFieldText(unsigned int fieldID, CString text);

    CString GetTMNote() const;
    void SetTMNote(CString note);
    CString GetPMNote() const;
    void SetPMNote(CString note);
    bool HasNote() const;
    bool HasTMNote() const;

    bool ContainsPercRectPoint(CPoint point);
    bool ContainsPercRectPoint(CPoint point, CRect& rc);
    void AddPercRect(CRect rect);
    bool ContainsPercWorkRectPoint(CPoint point);
    bool ContainsPercWorkRectPoint(CPoint point, CRect& rc);
    void AddPercWorkRect(CRect rect);
    bool ContainsActWorkRectPoint(CPoint point);
    bool ContainsActWorkRectPoint(CPoint point, CRect& rc);
    void AddActWorkRect(CRect rect);
	bool ContainsActStartRectPoint(CPoint point);
	bool ContainsActStartRectPoint(CPoint point, CRect& rc);
	void AddActStartRect(CRect rect);
	bool ContainsActFinishRectPoint(CPoint point);
	bool ContainsActFinishRectPoint(CPoint point, CRect& rc);
	void AddActFinishRect(CRect rect);
    bool ContainsActDurationRectPoint(CPoint point);
    bool ContainsActDurationRectPoint(CPoint point, CRect& rc);
    void AddActDurationRect(CRect rect);
    bool ContainsNotesRectPoint(CPoint point);
    bool ContainsNotesRectPoint(CPoint point, CRect& rc);
    void AddNotesRect(CRect rect);

    bool ContainsCustomText1RectPoint(CPoint point);
    bool ContainsCustomText1RectPoint(CPoint point, CRect& rc);
    void AddCustomText1Rect(CRect rect);
    bool ContainsCustomText2RectPoint(CPoint point);
    bool ContainsCustomText2RectPoint(CPoint point, CRect& rc);
    void AddCustomText2Rect(CRect rect);
    bool ContainsCustomText3RectPoint(CPoint point);
    bool ContainsCustomText3RectPoint(CPoint point, CRect& rc);
    void AddCustomText3Rect(CRect rect);
    bool ContainsCustomTextRectPoint(unsigned int fiekdID, CPoint point);
    bool ContainsCustomTextRectPoint(unsigned int fiekdID, CPoint point, CRect& rc);
    void AddCustomTextRect(unsigned int fiekdID, CRect rect);

    void ClearReportRects(); 

    long long GetTaskUID() const;
    void SetTaskUID(long long iAssgnUID);
    CString GetProjectName() const;
    void SetProjectName(CString prName);
    CString GetProjectGUID() const;
    void SetProjectGUID(CString prjGUID);
    long long GetTaskBaseUID() const;
    void SetTaskBaseUID(long long iBaseUID);
    CString GetTaskName() const;
    void SetTaskName(CString strTaskName);
    enumTUTaskWorkInsertMode GetWorkInsertMode() const;
    void SetWorkInsertMode(enumTUTaskWorkInsertMode eWorkInsertMode);
    bool IsUndefined();
    bool IsPercentCompleteActive();
    bool IsPercentWorkCompleteActive();
    bool IsActualWorkActive();
	bool IsActualStartActive();
	bool IsActualFinishActive();
    bool IsActualDurationActive();
    enumTUTaskReportsStatus GetStatus() const;
    void SetStatus(enumTUTaskReportsStatus eStatus);
    void SetSummary(bool isSummary);
    bool IsSummary() const;

    bool IsPercentCompleteEdited() const;    
    bool IsPercentWorkCompleteEdited() const;
    bool IsActualWorkEdited() const;         
	bool IsActualStartEdited() const;        
	bool IsActualFinishEdited() const;       
    bool IsActualDurationEdited() const;     
    bool IsNotesEdited() const;  
    bool IsCustomText1Edited() const;        
    bool IsCustomText2Edited() const;        
    bool IsCustomText3Edited() const;        
    bool IsCustomTextEdited(unsigned int fieldID) const;
    bool IsEdited(bool textField = true) const;                
    void SetPercentCompleteEdited(bool edited);    
    void SetPercentWorkCompleteEdited(bool edited);
    void SetActualWorkEdited(bool edited);         
	void SetActualStartEdited(bool edited);        
	void SetActualFinishEdited(bool edited);       
    void SetActualDurationEdited(bool edited);     
    void SetNotesEdited(bool edited);   
    void SetCustomText1Edited(bool edited);        
    void SetCustomText2Edited(bool edited);        
    void SetCustomText3Edited(bool edited);        
    void SetCustomTextEdited(unsigned int fieldID, bool edited);       
    bool IsSaved() const;                  
    void SetSaved(bool bSaved = true);  
    bool IsNotesSaved() const;                   
    void SetNotesSaved(bool bSaved = true);  
    bool IsCustomTextSaved() const;                   
    void SetCustomTextSaved(bool bSaved = true);      
    bool IsDisabled() const;                 
    void SetDisabled(bool bDisabled = true); 
    UINT_PTR GetUINT_PTR() const;            
    void SetUINT_PTR(UINT_PTR uint_ptr);     
    bool IsSelected() const;                 
    void SetSelected(bool bSel);             

    void EmptyReport(bool emtyTextFields = true);
    bool IsSaveable();

    unsigned int GetCustomText1FieldID() const;
    void SetCustomText1FieldID(unsigned int filedID);
    unsigned int GetCustomText2FieldID() const;
    void SetCustomText2FieldID(unsigned int filedID);
    unsigned int GetCustomText3FieldID() const;
    void SetCustomText3FieldID(unsigned int filedID);

    bool IsCustomTextField(unsigned int fieldID);
};
