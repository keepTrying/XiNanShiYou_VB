#ifdef  _WIN32
#define STDCALL  __stdcall
#else
#define STDCALL
#endif
#ifndef SDTAPI_
#define SDTAPI_
#ifdef __cplusplus
extern "C"{
#endif 

#pragma pack(1)
typedef struct tagIDCardData{
	char Name[32];
	char Sex[4];
	char Nation[6];
	char Born[18];
	char Address[72];
	char IDCardNo[38];
	char GrantDept[32];
	char UserLifeBegin[18];
	char UserLifeEnd[18];
	char reserved[38];
	char PhotoFileName[255];
}IDCardData;

#pragma pack()

/**********************************************************
 ********************** 端口类API *************************
 **********************************************************/
int STDCALL Syn_GetCOMBaud(int iComID,unsigned int *puiBaud);
int STDCALL Syn_SetCOMBaud(int iComID,unsigned int  uiCurrBaud,unsigned int  uiSetBaud);
int STDCALL Syn_OpenPort(int iPortID);
int STDCALL Syn_ClosePort(int iPortID);

/**********************************************************
 ********************** SAM类API **************************
 **********************************************************/
int STDCALL Syn_GetSAMStatus(int iPortID,int iIfOpen);
int STDCALL Syn_ResetSAM(int iPortID,int iIfOpen);
int STDCALL Syn_GetSAMID(int iPortID,unsigned char *pucSAMID,int iIfOpen);
int STDCALL Syn_GetSAMIDToStr(int iPortID,char *pcSAMID,int iIfOpen);

/**********************************************************
 ******************* 身份证卡类API ************************
 **********************************************************/
int STDCALL Syn_StartFindIDCard(int iPortID,unsigned char *pucManaInfo,int iIfOpen);
int STDCALL Syn_SelectIDCard(int iPortID,unsigned char *pucManaMsg,int iIfOpen);
int STDCALL Syn_ReadMsg(int iPortID,int iIfOpen,IDCardData *pIDCardData);

/**********************************************************
 ******************* 附加类API ************************
 **********************************************************/
int  STDCALL Syn_SendSound(int iCmdNo);
void STDCALL Syn_DelPhotoFile();

#ifdef __cplusplus
}
#endif 
#endif