module help_mod
contains
  subroutine help(Title)
  use nrtype, only:CR
  character(*),intent(in),optional::Title
  character(:),allocatable::Errmsg,Tit,Str
   Errmsg="ModelTest2XL [/pPathDir] [/mMarinCSVFile] [/sSIMAXLSXFile]"//CR
   Errmsg=Errmsg//'Eg. ModelTest2Xls /pLC01\node_1 /lLC100_01 /xB_Summary.xlsx /t"SIMA SIMAPlt" /sX /MWTF /MPM /M3hour /MBL=19748 /c1000.0 /n12'//CR
   Errmsg=Errmsg//'    ModelTest2Xls /pD:\SIMA_Run\CondSet\1-100 "/l3041Wd[.35,.4] CondSet_1" "/l3080Wd[.35,.4] CondSet_2" "/l3110Wd[.35,.4] CondSet_3" "/l3140Wd[.35,.4] CondSet_4" "/l3201Wd[.35,.4] CondSet_5" "/l3210Wd[.35,.4] CondSet_6" /xResSum.xlsx "/tSIMA_Data NumExtremeSum" /sX /MBL=16500 /MObs /M3hour /c1000.0 /n12'
  if(present(Title))then
     write(*,*)Title
  else
     write(*,*)ErrMsg
  end if
!   PBres=ProgDialog(0,Tit,.true.,Errmsg,IDC_OKOK,'Continue?'C,.true.,10); if(PBres==2)Stop
  stop
  end subroutine help
end module
    
Program ModelTest2XL
USE IFPORT; USE,INTRINSIC::ISO_FORTRAN_ENV,ONLY:IOSTAT_EOR
use util_xk,only:reallocate,CR  ,std 
use LStrFileUti
use libxl
IMPLICIT NONE
type(BookHandle)::Bukhdl
type(SheetHandle)::Arkhdl0,Arkhdl
type(FormatHandle)::Centerfmt,newFmt,fmt1,LTRB1fmt,LTRB2fmt
type(FontHandle)::font1
integer nfmt,nfont,iBukLoad,i_ret,rnum,cnum,iType,iMaxType
character(:),allocatable::crad,filnam,XLfil,SumXLFil,ArkNam0,ArkNam,TabHdr(:),StrArr(:)
!integer,parameter::MBL=10200
integer::iui=1,iflog=2,iuo=3,i,II,j,k,m,ios,nLC,mL,nL,nRec,nStep,nArg,iRow,iCol,nCol,nLine,MBL,iMea,iStd,iMin,iMax,j1(1),j2(1),nHr
CHARACTER(256)::buffer,str*20,arg,strL*500,str2,iomsg,n2s*2,DateStr*9
character(:),allocatable::path,iFil,oFil,VarList(:,:),FILOUT(:),MTBuk(:)!=['MTB','MTL']
real*8,allocatable::Table(:,:),LTen(:,:),LTmx(:)
real*8::pValue
integer,allocatable::ipos(:),iCol2Cal(:,:),iCo2Merg(:,:)
real::TransT
INTERFACE
subroutine IrrWaveTestSummary(Dir,iCol2Cal,iCo2Merg,XLfil,SumXLFil,TransT)
IMPLICIT  NONE
character(*),intent(in)::Dir,XLfil,SumXLFil
integer,allocatable::iCol2Cal(:,:),iCo2Merg(:,:)
real,intent(in)::TransT
end subroutine IrrWaveTestSummary
END INTERFACE
!allocate(character(64)::tmpStr)
!====== Turret Force Comparison vs Model Test Results show Inertia Corrected becomes larger than Numerical ========
!======  Turret Forces are taken as original without Modification. 24.02.2020

!call StatData('D:\WorkDir\ProjSpace\M2269_B29\model_test\data\wa3041.txt',0)
!allocate(LTmx(63536))
!  iFil='yPos.txt'; call FOpenFile(iui,iFil,'old');
!    do i = 1, size(LTmx)
!        read(iui, *, iostat=ios) LTmx(i)
!    end do
!
!pValue = std(LTMx)
!Path='X:\analysis\projects\2110_CaRongDo\model_test\timeseries\'!"D:\WorkDir\ProjSpace\M2269_B29\model_test\data\irrwave_timeseries\";
!Path='F:\source\For\ModelTest2Xls\'!"D:\WorkDir\ProjSpace\M2269_B29\model_test\data\irrwave_timeseries\";
!VarList=Reshape([&!'3041' is the 1st Ballast
!'3041','3011', &
!'3080','3020', &
!'3110','3030', &
!'3140','3050', &
!'3201','3060', &
!'3210','3070', &
!'3220','3090', &
!'','3130', &!'3090',&!'3070', &
!'','3150'],[nLC,2],order=[2,1]);
!SumXLFil='OSXMTemplate.xlsx'; !ArkNam0='Collection'; ArkNam='ExtremeSum';
!XLFil='OSXMTemplate.xlsx'; !ArkNam0='Collection'; ArkNam='ExtremeSum';
call CommandLine()
nLC=size(StrArr)/2
if(nLC>0)then
   VarList=Reshape(StrArr,[nLC,2],order=[2,1]);
   mL=nLC; nL=2
else
   VarList=Reshape(StrArr,[1,1]);    mL=1; nL=1
end if
!Path='D:\WorkDir\ProjSpace\M2269_B29\Sima\B29_MTB_CalibImpWvWi\CondSet\138-20241029130001\'
!nLC=4
!VarList=Reshape([&!'3041' is the 1st Ballast
!'3041','CondSet_1\', &
!'3080','CondSet_2\', &
!'3110','CondSet_3\', &
!'3140','CondSet_4\'],[nLC,2],order=[2,1]);
!
!!Path='D:\WorkDir\ProjSpace\M2269_B29\Sima\B29_MTL_CalibImpWvWi\ConditionSet\157-20241028211720\'
!Path='D:\WorkDir\ProjSpace\M2269_B29\Sima\B29_MTL_CalibImpWvWi\CondSet\157-20241029125844\'
!nLC=7
!VarList=Reshape([&!'3041' is the 1st Ballast
!'3011','CondSet_1\', &
!'3020','CondSet_2\', &
!'3030','CondSet_3\', &
!'3050','CondSet_4\', &
!'3060','CondSet_5\', &
!'3070','CondSet_6\', &
!'3090','CondSet_7\'],[nLC,2],order=[2,1]);
if(iType==-1)then
 do j=1,nL;  do i=1,mL
   ArkNam=VarList(i,j); m=clen(ArkNam)
   if(m>1)then
     if(ArkNam(m:m)=='\')m=m-1
     ArkNam=ArkNam(:m); XLfil=ctrim(Path)//ArkNam//'.xlsx'
     call WeiBullFitXLData(XLfil,ArkNam//''C,TransT)
   end if
 end do; end do
 stop
end if 
if(iType==1)then
  allocate(iCol2Cal(5,2),iCo2Merg(9,2))! Made for Nexus OSX vessel model test1339-MAT-W-RA-0001_00461684
  iCol2Cal(:,:)=Reshape([0,2,19,30,38,46,52,64,68,68],[5,2],order=[2,1]);
  iCo2Merg(:,:)=Reshape([1,2,3,14,15,19,20,21,22,23,24,26,27,33,34,36,37,37],[9,2],order=[2,1]);
 do j=1,nL;  do i=1,mL
   if(len_trim(VarList(i,j))>0)call IrrWaveTestSummary(Path,iCol2Cal,iCo2Merg,'Test'//ctrim(VarList(i,j))//'.sta.xlsx',XLFil,TransT);
 end do; end do
 stop
end if 
call reallocate(LTen,nLine,3); call reallocate(LTmx,nLine)
!XLfil='ModelTestResSum.xlsx';
!ArkNam0='Collection'; ArkNam='MTExtremeSum'; 
!ArkNam0='SIMA_Data'; ArkNam='NumExtremeSum'
Bukhdl=NEWXMLBook(); iBukLoad=xlBookLoad(Bukhdl,ctrim(XLfil)//""C); CALL DATE(DateStr)
IF(iBukLoad==1)THEN
  Arkhdl0=xlBookGetSheetbyName(Bukhdl,ArkNam0//''C)
  Arkhdl=xlBookGetSheetbyName(Bukhdl,ArkNam//''C)
  if(Arkhdl%Point==0)call ErrorMsg('Error in loading sheet:'//ArkNam//' of '//XLfil//CR//'Note: Case Sensitive for Names!',-1)
  iRow=0;j=0; call FindCell(Arkhdl,['Date: '],iRow,j,1,[20,1]); call WriteCell(Arkhdl,iRow,j+1,DateStr);
   iRow=4; i=iRow; i_ret=xlSheetCellType(Arkhdl,i,2);
    do while(i_ret==CELLTYPE_NUMBER)
     i=i+1; i_ret=xlSheetCellType(Arkhdl,i,2)
     if(i_ret==CELLTYPE_EMPTY.or.i_ret>=CELLTYPE_BLANK)then; iRow=i; exit; end if
    end do !CELLTYPE_EMPTY=0,CELLTYPE_NUMBER=1,CELLTYPE_STRING=2,CELLTYPE_BOOLEAN=3,CELLTYPE_BLANK=4,CELLTYPE_ERROR=5
    call WriteCell(Arkhdl0,iRow,(nLine+9)*4+1,DateStr//'|MaxType='//merge('WTF','Obs',iMaxType>0))
    call WriteCell(Arkhdl,iRow,17,DateStr//'|MaxType='//merge('WTF','Obs',iMaxType>0))
    i=xlBookSheetCount(Bukhdl)-1!NOTE 0-based sheet index
    call xlBookSetActiveSheet(Bukhdl,i)
ELSE
  Arkhdl0=xlBookAddSheet(Bukhdl,ArkNam0//""C)
  if(Arkhdl0%point/=0)then
   Centerfmt=AddBukFormat(Bukhdl,[2,1,0,0,0,0,1,1,1,1],3)
   LTRB1fmt=AddBukFormat(Bukhdl,[1,1,1,1],9);
   LTRB2fmt=AddBukFormat(Bukhdl,[6,6,6,6],9); 
   i_ret=xlSheetSetCol(Arkhdl0,0,1,dble(20)); iRow=2; 
   call WriteCell(Arkhdl0,0,0,'Date: '//DateStr,Centerfmt); !call WriteCell(Arkhdl0,1,0,[(dble(i),i=1,61)]);
    do j=0,84; 
     i_ret=xlSheetWriteFormula(Arkhdl0,1,j,'=COLUMN()'C); 
    end do
    i_ret=xlSheetSetMerge(Arkhdl0,iRow,iRow+1,0,0)!For existing excel file, MUST use xlSheetDelMerge(handle,row,col) first!
    call WriteCell(Arkhdl0,iRow,0,"LoadCase#_Seed",Centerfmt);     
    do i=1,nLine
     call WriteCell(Arkhdl0,iRow,(i-1)*4+1,"Mooring Line#"//num2str(i)); 
    end do
    i=nLine+1; call WriteCell(Arkhdl0,iRow,(i-1)*4+1,"FXY_TUR [kN]");
    i=nLine+2; call WriteCell(Arkhdl0,iRow,(i-1)*4+1,"FZ_TUR [kN]");
    i=nLine+3; call WriteCell(Arkhdl0,iRow,(i-1)*4+1,"XY_TUR [m]");
    i=nLine+4; call WriteCell(Arkhdl0,iRow,(i-1)*4+1,"X_Pos [m]");
    i=nLine+5; call WriteCell(Arkhdl0,iRow,(i-1)*4+1,"Y_Pos [m]"C);
    i=nLine+6; call WriteCell(Arkhdl0,iRow,(i-1)*4+1,"Z_Pos [m]");
    i=nLine+7; call WriteCell(Arkhdl0,iRow,(i-1)*4+1,"Roll ["//char(176)//"]");
    i=nLine+8; call WriteCell(Arkhdl0,iRow,(i-1)*4+1,"Pitch ["//char(176)//"]");
    i=nLine+9; call WriteCell(Arkhdl0,iRow,(i-1)*4+1,"YAW ["//char(176)//"]"C);
    do i=1,(nLine+9)*4,4
     i_ret=xlSheetSetMerge(Arkhdl0,iRow,iRow,i,i+3);   
    call SetCellFormat(Bukhdl,Arkhdl0,iRow,i,-iRow,-(i+3),LTRB1fmt,'LTRB');
    end do  
    TabHdr=[" ",(("Mean","Std","Min","Max"),i=1,nLine+9)]; 
    call WriteCell(Arkhdl0,iRow+1,0,TabHdr,Centerfmt);
    call WriteCell(Arkhdl0,iRow,(nLine+9)*4+1,DateStr//'|MaxType='//merge('WTF','Obs',iMaxType>0))
  end if
  Arkhdl=xlBookAddSheet(Bukhdl,ArkNam//""C)
  if(Arkhdl%point/=0)then
   i_ret=xlSheetSetCol(Arkhdl,0,1,dble(20)); iRow=2; 
   call WriteCell(Arkhdl,0,0,'Date:'//DateStr,Centerfmt); call WriteCell(Arkhdl,0,1,DateStr,Centerfmt); !call WriteCell(Arkhdl,1,0,[(dble(i),i=1,17)]);
   do j=0,64; i_ret=xlSheetWriteFormula(Arkhdl0,1,j,'=COLUMN()'C); end do
    i_ret=xlSheetSetMerge(Arkhdl,iRow,iRow+1,0,0)
    i_ret=xlSheetSetMerge(Arkhdl,iRow,iRow,1 ,5 ); i_ret=xlSheetSetMerge(Arkhdl,iRow,iRow,6 ,8 ); i_ret=xlSheetSetMerge(Arkhdl,iRow,iRow,9 ,10);
    i_ret=xlSheetSetMerge(Arkhdl,iRow,iRow,11,12); i_ret=xlSheetSetMerge(Arkhdl,iRow,iRow,13,14); i_ret=xlSheetSetMerge(Arkhdl,iRow,iRow,15,16);
!    TabHdr=["LoadCase#_Seed","WorstLine: kN|kN|kN|-|#","2nd Worst: kN|kN|#","xy[m]","Heading["//char(194)//char(176)//"C]"]; 
    call WriteCell(Arkhdl,iRow,0,"LoadCase#_Seed",Centerfmt);     call WriteCell(Arkhdl,iRow,1,"WorstLine: kN|kN|kN|-|#"); 
    call WriteCell(Arkhdl,iRow,6,"2nd Worst: kN|kN|#"); call WriteCell(Arkhdl,iRow,9,"Fxy kN"); call WriteCell(Arkhdl,iRow,11,"Fz kN"); call WriteCell(Arkhdl,iRow,13,"XY m");
    call WriteCell(Arkhdl,iRow,15,"Heading["//char(176)//"]"C);
    call SetCellFormat(Bukhdl,Arkhdl,iRow,6,-iRow,-8,LTRB1fmt,'LR'); call SetCellFormat(Bukhdl,Arkhdl,iRow,11,-iRow,-12,LTRB1fmt,'LR');
    call SetCellFormat(Bukhdl,Arkhdl,iRow,13,-iRow,-14,LTRB1fmt,'R'); !call SetCellFormat(Bukhdl,Arkhdl,iRow,15,-iRow,-16,LTRB1fmt,'R');
    TabHdr=[" ","Mean","Max","Std","Saf","L#","Mean","Max","L#","Mean","Max","Mean","Max","Mean","Max","Mean","Std"]; 
    call WriteCell(Arkhdl,iRow+1,0,TabHdr,Centerfmt)
    !call WriteCell(Arkhdl,iRow+3,0,[(dble(i),i=1,15)],rowcol='c');
    call SetCellFormat(Bukhdl,Arkhdl,iRow,0,-(iRow+100),-16,LTRB2fmt,'LTR');
    call xlBookSetActiveSheet(Bukhdl,1)
    i_ret=xlBookSave(Bukhdl, XLfil//""C)
  end if
  iRow=iRow+2
END IF
  
select case(iType)
case(2)
  write(*,'(a)')'Processing... SIMO/A simulation bin-XLS results: Path=',Path,'Xlfil=',XLfil,'Tabs=',ArkNam0,ArkNam
  call SIMAXLSData(VarList,Path,XLfil,ArkNam0,ArkNam,Arkhdl0,Arkhdl,TransT,iMaxType,nHr)
case(3)
  write(*,'(a)')'Processing... Sintef XLS Model Test Data Extraction/Stats/MaxObs or MaxWeibull fitted: ',Path,XLfil,ArkNam
  call SIOModelTestXLSData(Path,VarList,Arkhdl0,XLfil,Arkhdl,TransT,iMaxType)  
case(4)
  write(*,'(a)')'Processing... Marin CSV Model Test Data Extraction/Stats/Weibul fitting: ',Path,XLfil,ArkNam
  call MarinModelTestCSVData(VarList)
end select
  
call xlBookRelease(Bukhdl); 

contains

subroutine MarinModelTestCSVData(VarList)
character(:),allocatable::VarList(:,:)
character(:),allocatable::StrArr(:)
real*8,allocatable::NumArr(:)
integer::nStr,nNum
open(iflog,file='logcsv.txt')
DO II=1,nLC
  iFil=Path//trim(VarList(II,1))//trim(VarList(II,2))//"_UF.csv" !31448_01OB_03_ 003_001_01_UF.csv
  !trim(VarList(II,3))
 if(FILE_EXIST(iFil))then
  open(iui,file=iFil,status='old'); ios=0; ncol=0; iMea=0
  ! ios=lineread(iui,crad,ncol);
   do while(ios/=-1)      
      ios=lineread(iui,FILOUT,ipos); if(ios==-1)exit
      call LineParser(FILOUT,ipos,strArr,numArr)
      nStr=size(strArr); nNum=size(numArr)
      if(nNum==0.and.iMea==0)then
        iMea=(index(strArr,'Mean')+1)/2
        iStd=(index(strArr,'StdDev')+1)/2
        iMin=(index(strArr,'Min')+1)/2
        iMax=(index(strArr,'Max')+1)/2
      elseif(nStr>=nNum)then
       if(nNum>=nLine)then
        call reallocate(Table,II,64)
        if(instr(strArr(1),'F_MOOR'))then
         do k=1,nLine
          if(StrEqi(strArr(1),'F_MOOR.0'//char(48+k)//'_P'))exit
         end do; if(k>nLine)call errormsg(iFil//' contains None Mooring loads!',-1)
         Table(II,(k-1)*4+1:k*4)=numArr([iMea,iStd,iMin,iMax])! numArr([1,2,7,6])
         LTen(k,:)=numArr([iMea,iMax,iStd])! numArr([1,6,2])
        elseif(StrEqi(strArr(1),'FX_TUR'))then
         !Table(II,37:40)=numArr([iMea,iStd,iMin,iMax])
         Table(II,4*nLine+1:4*(nLine+1))=numArr([iMea,iStd,iMin,iMax])
        elseif(StrEqi(strArr(1),'FY_TUR'))then
         !Table(II,41:44)=numArr([iMea,iStd,iMin,iMax])
         Table(II,4*(nLine+1)+1:4*(nLine+2))=numArr([iMea,iStd,iMin,iMax])
        elseif(StrEqi(strArr(1),'FZ_TUR'))then
!         Table(II,45:48)=numArr([iMea,iStd,iMin,iMax])
         Table(II,4*(nLine+2)+1:4*(nLine+3))=numArr([iMea,iStd,iMin,iMax])            
        elseif(StrEqi(strArr(1),'XY_TUR'))then
!         Table(II,49:52)=numArr([iMea,iStd,iMin,iMax])
         Table(II,4*(nLine+3)+1:4*(nLine+4))=numArr([iMea,iStd,iMin,iMax])            
        elseif(StrEqi(strArr(1),'YAW'))then
!         Table(II,53:56)=numArr([iMea,iStd,iMin,iMax])
         Table(II,4*(nLine+4)+1:4*(nLine+5))=numArr([iMea,iStd,iMin,iMax])            
        elseif(StrEqi(strArr(1),'X_TUR_SWIV'))then
!         Table(II,57:60)=numArr([iMea,iStd,iMin,iMax])
         Table(II,4*(nLine+5)+1:4*(nLine+6))=numArr([iMea,iStd,iMin,iMax])
        elseif(StrEqi(strArr(1),'Y_TUR_SWIV'))then
!         Table(II,61:64)=numArr([iMea,iStd,iMin,iMax])
         Table(II,4*(nLine+6)+1:4*(nLine+7))=numArr([iMea,iStd,iMin,iMax])            
        end if
       end if
         ! nLines=int(numArr(1))
      else
       call errormsg(iFil//' contains invalid input numbers!',-1)
      end if          
   end do
   
   call WriteCell(Arkhdl0,iRow,0,trim(VarList(II,2))//":"//trim(VarList(II,3))//""C);
   call WriteCell(Arkhdl0,iRow,1,Table(II,:));
   call WriteCell(Arkhdl,iRow,0,trim(VarList(II,2))//":"//trim(VarList(II,3))//""C);
   LTmx=LTen(:,2); j1=maxloc(LTmx(:)); LTmx(j1)=0d0; j2=maxloc(LTmx(:));
!   call WriteCell(Arkhdl,iRow,1,[LTen(j1,:),MBL/LTen(j1,2),dble(j1),LTen(j2,1:2),dble(j2),Table(II,37:40:3),Table(II,41:44:3),Table(II,45:47:2),Table(II,49:52:3),Table(II,53:54)]);
   call WriteCell(Arkhdl,iRow,1,[LTen(j1,:),MBL/LTen(j1,2),dble(j1),LTen(j2,1:2),dble(j2),Table(II,4*nLine+1:4*(nLine+1):3),&
       Table(II,4*(nLine+1)+1:4*(nLine+2):3),Table(II,4*(nLine+2)+1:4*(nLine+3)-1:2),Table(II,4*(nLine+3)+1:4*(nLine+4):3),Table(II,4*(nLine+4)+1:4*(nLine+4)+2)]);
   i_ret=xlBookSave(Bukhdl, XLfil//""C); iRow=iRow+1
  
  close(iui); write(iflog,'(a)')iFil//' has been read successfully!'
 else
  write(iflog,'(a)')"File://"//iFil//' does not EXIST!'
 end if
END DO; close(iflog)
End subroutine MarinModelTestCSVData

subroutine SIOModelTestXLSData(Path,VarList,Arkhdl0,XLfil,Arkhdl,TransT,iMaxType)
use UTIL_xk,only:Mean,Std,ErrorMsg
character(:),allocatable::Path,XLfil
character(:),allocatable::VarList(:,:)
type(SheetHandle),intent(in)::Arkhdl0,Arkhdl
real,intent(in)::TransT
integer,intent(in)::iMaxType
type(SheetHandle)::hArk
type(BookHandle)::hBuk
character(:),allocatable::Shnam,filnam,Rend,tempFile,ErrMsg
integer::i,j,k,II,iR,jC,iLS,jLS,nCol,i_ret,iR0,k_end,nTr,nRec
integer,allocatable::iCol2Cal(:,:)
real*8,allocatable::Table(:),ColData(:)
real*8::MeanStdMin(3),dum,tStep
logical::isFz
iCol2Cal=Reshape([3, 14, 22,22,17,17,33,33,27,32,36,36],[6,2],order=[2,1]);
!                 L1 L12 Fxy   Fz    Rxy   XYZposRollPitchYaw ZOG
DO jLS=1,size(VarList,2)! Ballast & Loaded
  DO iLS=1,size(VarList,1)! nLC
    IF(len_trim(VarList(iLS,jLS))>0)THEN
      hBuk=NEWXMLBook(); Shnam=ctrim(VarList(iLS,jLS)); k=clen(Shnam)
      if(Shnam(k:k)=='\')Shnam=Shnam(:k-1)
      filnam=ctrim(Path)//Shnam
      iBukload=xlBookLoad(hBuk,filnam//'.xlsx'C)
      if(iBukload==0)then
         call ErrorMsg('Error in loading sheet:'//filnam//'.xlsx',5)
      end if    
      hArk=xlBookGetSheetbyName(hBuk,Shnam//''C); 
      IF(hArk%Point>0)then
      write(*,'(a)')'Now processing: '//filnam//'.xlsx, sheet:'//Shnam
        iR0=1; i_ret=xlSheetCellType(hArk,iR0,0);
        do while(i_ret/=CELLTYPE_NUMBER)
         iR0=iR0+1; ; i_ret=xlSheetCellType(hArk,iR0,0)
        end do; iR=iR0
        ! Write initial X,Y,Z and Yaw condition
        do i=27,32
          call WriteCell(Arkhdl,iRow,i,xlSheetReadNum(hArk,iR0,i))
        end do
        do while(i_ret==CELLTYPE_NUMBER)
         iR=iR+1; i_ret=xlSheetCellType(hArk,iR,0)
         if(i_ret==CELLTYPE_STRING.or.i_ret==CELLTYPE_EMPTY.or.i_ret>=CELLTYPE_BLANK)then; 
            Rend=num2str(iR);       !1-based Number = iR-th Row#
            exit;                   !0-based Row Number = iR which is a string cell =>'Mean'
         end if
        end do !CELLTYPE_EMPTY=0,CELLTYPE_NUMBER=1,CELLTYPE_STRING=2,CELLTYPE_BOOLEAN=3,CELLTYPE_BLANK=4,CELLTYPE_ERROR=5
        !Find the time step from two rows of the 1st Column jc=0
        tStep=(xlSheetReadNum(hArk,iR-1,0)-xlSheetReadNum(hArk,iR0,0))/real(iR-iR0-1)
        nTr=int(TransT/tStep); nRec=iR-iR0-nTr
        call reallocate(ColData,nRec)
        call reAllocate(Table,4*sum(iCol2Cal(:,2)-iCol2Cal(:,1)+1))
        !do II=1,size(iCol2Cal,1)
        II=1! For Mooring Lines
        do jc=iCol2Cal(II,1),iCol2Cal(II,2)
          do i=1,nRec; k=i+iR0+nTr-1
            ColData(i)=xlSheetReadNum(hArk,k,jc)
            ErrMsg=GetErrorMsg(hBuk);
            if(ErrMsg/='ok')call ErrorMsg(ErrMsg//' @('//num2str(k)//','//num2str(jc)//') in file:'//filnam//'.xlsx',10)
          end do  
          MeanStdMin=[mean(ColData),std(ColData),minval(ColData)]
          j=jc-iCol2Cal(II,1)+1! For the Mooring Line# j
         do i=0,2
          dum=xlSheetReadNum(hArk,iR+i,jc)
          if(abs(dum)<1d-5)then
           Table((j-1)*4+i+1)=MeanStdMin(i+1)
          else
           Table((j-1)*4+i+1)=dum
          end if    
         end do; i=3
         Table((j-1)*4+i+1)=xlSheetReadNum(hArk,iR+i+iMaxType,jc);!iR+3 is ObsMax, iR+4 is Max_WeibulTailFitted
         LTen(j,:)=Table([(j-1)*4+1,j*4,(j-1)*4+2])! Mean, MaxT, Std
        end do; k_end=j*4
        DO II=2,size(iCol2Cal,1)
         Do jc=iCol2Cal(II,1),iCol2Cal(II,2)
          do i=1,nRec; k=i+iR0+nTr-1
            ColData(i)=xlSheetReadNum(hArk,k,jc)
            ErrMsg=GetErrorMsg(hBuk);
            if(ErrMsg/='ok')call ErrorMsg(ErrMsg//' @('//num2str(k)//','//num2str(jc)//') in file:'//filnam//'.xlsx',10)
          end do; dum=maxval(ColData) 
          isFz=dum<0.or.abs(dum)<abs(minval(ColData))
          if(isFz)then
              MeanStdMin=[mean(ColData),std(ColData),maxval(ColData)]
          else    
              MeanStdMin=[mean(ColData),std(ColData),minval(ColData)]
          end if    
           j=jc-iCol2Cal(II,1)+1
          
          do i=0,2
           if(isFz.and.i==2)then
              dum=xlSheetReadNum(hArk,iR+i+1,jc)! read the maximum as minimum amplitude
           else   
              dum=xlSheetReadNum(hArk,iR+i,jc)
           end if
           if(abs(dum)<1d-5)then
            Table(k_end+(j-1)*4+i+1)=MeanStdMin(i+1)
           else
            Table(k_end+(j-1)*4+i+1)=dum
           end if    
          end do; 
          if(iMaxType==0.and.isFz)then! If Observed Maximum and Negative f.eg: Fz, then take Observed Minimum
            i=2
          else
            i=3
          end if
          Table(k_end+(j-1)*4+4)=xlSheetReadNum(hArk,iR+i+iMaxType,jc);!iR+3 is ObsMax, iR+4 is Max_WeibulTailFitted
         END do; k_end=k_end+j*4
        END DO
      call xlBookRelease(hBuk);         
      !END if
      call WriteCell(Arkhdl0,iRow,0,'Test'//Shnam);
      call WriteCell(Arkhdl,iRow,0,'Test'//Shnam);
      call WriteCell(Arkhdl0,iRow,1,Table(:));
      LTmx=LTen(:,2); j1=maxloc(LTmx(:)); LTmx(j1)=0d0; j2=maxloc(LTmx(:)); 
      call WriteCell(Arkhdl,iRow,1,[LTen(j1,:),MBL/LTen(j1,2),dble(j1),LTen(j2,1:2),dble(j2),(Table(4*nLine+1+4*(i-1):4*nLine+1+4*i:3),i=1,3)]);
      i_ret=xlBookSave(Bukhdl,XLfil//""C); iRow=iRow+1
     END IF  
    END IF  
  END do
END DO


End subroutine SIOModelTestXLSData

subroutine SIMAXLSData(VarList,Path,XLfil,ArkNam0,ArkNam,Arkhdl0,Arkhdl,TransT,iMaxType,nHr)
use UTIL_xk,only:Mean,Std,ErrorMsg
use StatPack,only:WeiBulTailExtreme
character(:),allocatable::VarList(:,:)
type(BookHandle)::hBuk
type(SheetHandle)::hArk,Arkhdl0,Arkhdl
real::TransT,Dt
integer,intent(in)::iMaxType,nHr
character(:),allocatable::Path,XLfil,ArkNam0,ArkNam,XLFileList(:)
character(:),allocatable::Shnam,filnam,Rend,tmpstr,ErrMsg
integer::i,j,k,iR,jC,iLS,jLS,nCol,i_ret,iR0,k_end,nLin,iXL,nTr,mL,nL,k0,k1,nTot,nT_3h
integer,allocatable::Col_Elmfor(:),Col_Noddis(:),Col_result(:)
real*8,allocatable::ColData(:,:),DumArr(:)
!real*8,allocatable::Table(:),ColData(:,:),DumArr(:)
real*8,allocatable::Table(:,:),LTen(:,:,:)
real*8::MeanStdMinMax(4),dum
! Stats
integer,parameter::nMom=4,nCLS=10; 
real*8,DIMENSION(:),allocatable::X,XParent,XMaxima,XMinima,Pr,pdf,CDF
real*8::LOC_OLD,SCA_OLD,SHA_OLD,PLKey,P_ext,CLsz,gama,beta,mu,xMomc(0:nMom),GaBeMu(3),MLEpar(1),MMEPar(2),StatPar(6)
character(:),allocatable::fName,xTitle,yTitle
integer::pUnique,ErrType,iSign,n3Hr
!pValue: Statistic percentile=>any one of the hundred groups so divided while fractile is the value of a distribution for which some fraction of the sample lies below.
real*8::pValue=.37d0,pLow=.80d0,Twind,xExt(2),WeibPara(3)!=PFRAC; else; pValue=.37d0
!if(pres_PLKey)then; pLow=PLKey; else; pLow=.80d0; endif
! L1 L12 Fxy   Fz    Rxy   Xpos  Ypos Zpos Roll Pitch YAW
!Col_Elmfor=[4,5,6,7,8,9,10,11,12,13,14,15, 2, 3, 1];
Col_Elmfor=[(3+i,i=1,nLine), 2, 3, 1];
!           L1...                      L12 Fx Fy Fz    !Rxy   Xpos  Ypos  YAW
Col_Noddis=[1,   2,   3];
!           XPos YPos ZPos
Col_result=[4,   5,    6];
!           Roll Pitch Yaw
mL=size(VarList,1); nL=size(VarList,2)
!call reAllocate(Table,4*(size(Col_Elmfor)+size(Col_Noddis)+size(Col_result)))!call reallocate(Table,iR-2,nCol2Use)
  DO iLS=1,mL
    !IF(clen(VarList(iLS,nL))>0)THEN
      if(allocated(XLFileList))deallocate(XLFileList)
      i=DirFindFile(ctrim(Path//VarList(iLS,nL)),XLFileList,'.xlsx')
      if(i/=0)then
        i=DirFindFile(ctrim(Path),XLFileList,'.xlsx')  
        if(i/=0)then
          call ErrorMsg('No valid EXCEL file exists in: '//Path//' or '//VarList(iLS,nL),-1)
        else; iXL=Index_instr(XLFileList,'elmfor')
        end if
      else
        iXL=Index_instr(XLFileList,'elmfor')
      end if    
     write(*,'(a)')'Now processing: '//XLFileList(iXL)
      hBuk=NEWXMLBook(); iBukload=xlBookLoad(hBuk,XLFileList(iXL)//''C)
      !iBukload=xlBookLoadPartiallyUsingTempFile(hBuk,Path//VarList(iLS,nL)//'key_sima_elmfor.txt.xlsx'C,0,0,125904,'temp',1)
      ErrMsg=GetErrorMsg(hBuk);
        if(ErrMsg/='ok')call ErrorMsg('Error in loading sheet:'//XLFileList(iXL),5)
      hArk=xlBookGetSheetbyName(hBuk,"TimeSeries"C); nCol=size(Col_Elmfor)
      !IF(hArk%PTR)then
        iR0=1; i_ret=xlSheetCellType(hArk,iR0,0);
        do while(i_ret/=CELLTYPE_NUMBER.and.iR0<=512)
         iR0=iR0+1; ; i_ret=xlSheetCellType(hArk,iR0,0)
        end do; iR=iR0
        if(iR0>512)call ErrorMsg('Not complete data loading sheet:'//XLFileList(iXL),-5)
        do while(i_ret==CELLTYPE_NUMBER)
         iR=iR+1; i_ret=xlSheetCellType(hArk,iR,0)
         if(i_ret==CELLTYPE_STRING.or.i_ret==CELLTYPE_EMPTY.or.i_ret>=CELLTYPE_BLANK)then; 
            Rend=num2str(iR);       !1-based Number = iR-th Row#
            exit;                   !0-based Row Number = iR which is a string cell =>'Mean'
         end if
        end do !CELLTYPE_EMPTY=0,CELLTYPE_NUMBER=1,CELLTYPE_STRING=2,CELLTYPE_BOOLEAN=3,CELLTYPE_BLANK=4,CELLTYPE_ERROR=5
        nTot=iR-iR0; call reallocate(ColData,nTot,nCol); call reallocate(DumArr,nTot)
        do i=iR0,iR-1
           DumArr(i-iR0+1)=xlSheetReadNum(hArk,i,0)! Note 0-based indices           
           do II=1,nCol
            ColData(i-iR0+1,II)=xlSheetReadNum(hArk,i,Col_Elmfor(II))
           end do
        end do; Twind=DumArr(nTot)-DumArr(1)-TransT; dt=DumArr(2)-DumArr(1); nTr=int(TransT/dt)+1;
        ErrMsg=GetErrorMsg(hBuk);
        if(ErrMsg/='ok')call ErrorMsg('Error in reading sheet:'//XLFileList(iXL),5)
        nLin=count(Col_Elmfor>=Col_Elmfor(1))
        if(nLin/=nLine)call ErrorMsg('Error in Mooring Line Number input to SIMAXLSData',-1)

! If nHr is not given Defaulted to 0, then the maximum is to take as the whole time series, so #of 3hour maxima n3Hr=1
n3Hr=1; nT_3h=nTot
If(nHr>0)then
  Twind=nHr*3600.
  n3Hr=int(DumArr(nTot)/Twind)
  nT_3h=int(Twind/dt)
end if
call reallocate(LTen,n3Hr,nLine,3)
call reAllocate(Table,n3Hr,4*(size(Col_Elmfor)+size(Col_Noddis)+size(Col_result)))
       DO k=1,n3Hr
          k0=nTr+nT_3h*(k-1); k1=min(k0+nT_3h-1,nTot)
          DO j=1,nLine
            MeanStdMinMax(:3)=[mean(ColData(k0:k1,j)),std(ColData(k0:k1,j)),minval(ColData(k0:k1,j))]
            do i=0,2
            !dum=xlSheetReadNum(hArk,iR+i,Col_Elmfor(j)) this was used Mean, Std and Min rows for the whole timeseries processed by SIMAPP
            !if(abs(dum)<1d-5)then
             Table(k,(j-1)*4+i+1)=MeanStdMinMax(i+1)
            !else
            ! Table(k,(j-1)*4+i+1)=dum
            !end if
            end do; i=3
!           Table((j-1)*4+i+1)=xlSheetReadNum(hArk,iR+i+1,jc-1);!iR+3 is ObsMax, iR+4 is Max_WeibulTailFitted
            if(iMaxType>0)then
                call WeiBulTailExtreme(ColData(k0:k1,j),Twind,xExt,pValue,WeibPara)!,'MLine#'//num2str(j)
                Table(k,(j-1)*4+i+1)=xExt(1)
            else    
                Table(k,(j-1)*4+i+1)=maxval(ColData(k0:k1,j))
            end if    
            LTen(k,j,:)=Table(k,[(j-1)*4+1,j*4,(j-1)*4+2])! Mean, MaxT,Std
            k_end=j*4
          END DO
          !call xlBookRelease(hBuk)
!        !Note: DumArr will be reallocated to have smaller size starting with nTr:
        DumArr=sqrt(ColData(k0:k1,nLin+1)**2+ColData(k0:k1,nLine+2)**2)!Fxy=sqrt(Fx^2+Fy^2)
        if(iMaxType>0)then
            call WeiBulTailExtreme(DumArr,Twind,xExt,pValue,WeibPara)!,'Fxy=sqrt(Fx^2+Fy^2)'
            MeanStdMinMax=[mean(DumArr),std(DumArr),minval(DumArr),xExt(1)]
        else
            MeanStdMinMax=[mean(DumArr),std(DumArr),minval(DumArr),maxval(DumArr)]
        end if
        forall(i=1:4)Table(k,nLin*4+i)=MeanStdMinMax(i)

        dum=maxval(ColData(k0:k1,nCol))
        if(dum<0.or.abs(dum)<abs(minval(ColData(k0:k1,nCol))))then; iSign=-1; else; iSign=1; endif
        if(iMaxType>0)then
            call WeiBulTailExtreme(iSign*ColData(k0:k1,nCol),Twind,xExt,pValue,WeibPara);!,'Fz'
            MeanStdMinMax=[mean(ColData(k0:k1,nCol)),std(ColData(k0:k1,nCol)),minval(ColData(k0:k1,nCol)),iSign*xExt(1)]!Fz
        else    
            MeanStdMinMax=[mean(ColData(k0:k1,nCol)),std(ColData(k0:k1,nCol)),minval(ColData(k0:k1,nCol)),iSign*maxval(iSign*ColData(k0:k1,nCol))]!Fz
        end if
        forall(i=1:4)Table(k,nLin*4+4+i)=MeanStdMinMax(i)
       END DO;    
!Next is for noddis xlsx file
      iXL=Index_instr(XLFileList,'noddis');
     write(*,'(a)')'Now processing: '//XLFileList(iXL)      
      hBuk=NEWXMLBook();  iBukload=xlBookLoad(hBuk,XLFileList(iXL)//''C)
      if(iBukload==0)then
        if(ErrMsg/='ok')call ErrorMsg('Error in loading sheet:'//XLFileList(iXL),5)
      end if    
      hArk=xlBookGetSheetbyName(hBuk,"TimeSeries"C); nCol=size(Col_Noddis)
      !IF(hArk%PTR)then
        iR0=1; i_ret=xlSheetCellType(hArk,iR0,0);
        do while(i_ret/=CELLTYPE_NUMBER)
         iR0=iR0+1; ; i_ret=xlSheetCellType(hArk,iR0,0)
        end do; iR=iR0
        do while(i_ret==CELLTYPE_NUMBER)
         iR=iR+1; i_ret=xlSheetCellType(hArk,iR,0)
         if(i_ret==CELLTYPE_STRING.or.i_ret==CELLTYPE_EMPTY.or.i_ret>=CELLTYPE_BLANK)then; 
            Rend=num2str(iR);       !1-based Number = iR-th Row#
            exit;                   !0-based Row Number = iR which is a string cell =>'Mean'
         end if
        end do !CELLTYPE_EMPTY=0,CELLTYPE_NUMBER=1,CELLTYPE_STRING=2,CELLTYPE_BOOLEAN=3,CELLTYPE_BLANK=4,CELLTYPE_ERROR=5
        nTot=iR-iR0; call reallocate(ColData,nTot,nCol); call reallocate(DumArr,nTot)
        do i=iR0,iR-1
           DumArr(i-iR0+1)=xlSheetReadNum(hArk,i,0)! for 0-based indices           
           do II=1,nCol
            ColData(i-iR0+1,II)=xlSheetReadNum(hArk,i,Col_Noddis(II))
           end do
        end do;
        Twind=DumArr(nTot)-DumArr(1)-TransT; dt=DumArr(2)-DumArr(1); nTr=int(TransT/dt)+1;
If(nHr>0)then
  Twind=nHr*3600.
  n3Hr=int(DumArr(nTot)/Twind)
  nT_3h=int(Twind/dt)
end if
      ! Note: DumArr will be reallocated to have smaller size starting with k0:k1
       DO k=1,n3Hr
          k0=nTr+nT_3h*(k-1); k1=min(k0+nT_3h-1,nTot)
          DumArr=sqrt(ColData(k0:k1,1)**2+ColData(k0:k1,2)**2)!Rxy=sqrt(XPos^2+YPos^2)
          if(iMaxType>0)then
            call WeiBulTailExtreme(DumArr,Twind,xExt,pValue,WeibPara)!'Rxy=sqrt(XPos^2+YPos^2)'
            MeanStdMinMax=[mean(DumArr),std(DumArr),minval(DumArr),xExt(1)]
          else  
            MeanStdMinMax=[mean(DumArr),std(DumArr),minval(DumArr),maxval(DumArr)]
          end if
          forall(i=1:4)Table(k,(nLin+2)*4+i)=MeanStdMinMax(i)
          
          if(iMaxType>0)then
          if(maxval(ColData(k0:k1,1))<0)then; iSign=-1; else; iSign=1; endif
            call WeiBulTailExtreme(iSign*ColData(k0:k1,1),Twind,xExt,pValue,WeibPara)! XPos @Col.B=Turret2.D1Axial.S1
            MeanStdMinMax=[mean(ColData(k0:k1,1)),std(ColData(k0:k1,1)),minval(ColData(k0:k1,1)),iSign*xExt(1)]
          else  
            MeanStdMinMax=[mean(ColData(k0:k1,1)),std(ColData(k0:k1,1)),minval(ColData(k0:k1,1)),iSign*maxval(iSign*ColData(k0:k1,1))]
          end if  
          forall(i=1:4)Table(k,(nLin+3)*4+i)=MeanStdMinMax(i)
          
          if(maxval(ColData(k0:k1,2))<0)then; iSign=-1; else; iSign=1; endif
          if(iMaxType>0)then
            call WeiBulTailExtreme(iSign*ColData(k0:k1,2),Twind,xExt,pValue,WeibPara)! YPos
            MeanStdMinMax=[mean(ColData(k0:k1,2)),std(ColData(k0:k1,2)),minval(ColData(k0:k1,2)),iSign*xExt(1)]
          else  
            MeanStdMinMax=[mean(ColData(k0:k1,2)),std(ColData(k0:k1,2)),minval(ColData(k0:k1,2)),iSign*maxval(iSign*ColData(k0:k1,2))]
          end if  
          forall(i=1:4)Table(k,(nLin+4)*4+i)=MeanStdMinMax(i)
          
          do j=1,size(Col_Noddis)-2
           if(maxval(ColData(k0:k1,2+j))<0)then; iSign=-1; else; iSign=1; endif; yTitle=ReadCellStr(hArk,iR0-1,2+j)
           if(iMaxType>0)then
             call WeiBulTailExtreme(iSign*ColData(k0:k1,j+2),Twind,xExt,pValue,WeibPara)!,yTitle
             MeanStdMinMax=[mean(ColData(k0:k1,j+2)),std(ColData(k0:k1,j+2)),minval(ColData(k0:k1,j+2)),iSign*xExt(1)]
           else  
             MeanStdMinMax=[mean(ColData(k0:k1,j+2)),std(ColData(k0:k1,j+2)),minval(ColData(k0:k1,j+2)),iSign*maxval(iSign*ColData(k0:k1,j+2))]
           end if
           forall(i=1:4)Table(k,(nLin+4+j)*4+i)=MeanStdMinMax(i)
          end do
       END DO
        ErrMsg=GetErrorMsg(hBuk); call xlBookRelease(hBuk)
        if(ErrMsg/='ok')call ErrorMsg('Error in reading sheet:'//XLFileList(iXL),5)
!Next is for results xlsx file
      iXL=Index_instr(XLFileList,'results');
     write(*,'(a)')'Now processing: '//XLFileList(iXL)
      hBuk=NEWXMLBook();  iBukload=xlBookLoad(hBuk,XLFileList(iXL)//''C)
      if(iBukload==0)then
         call ErrorMsg('Error in loading sheet:'//XLFileList(iXL),5)
      end if    
      hArk=xlBookGetSheetbyName(hBuk,"TimeSeries"C); nCol=size(Col_Result)
      !IF(hArk%PTR)then
        iR0=1; i_ret=xlSheetCellType(hArk,iR0,0);
        do while(i_ret/=CELLTYPE_NUMBER)
         iR0=iR0+1; ; i_ret=xlSheetCellType(hArk,iR0,0)
        end do; iR=iR0
        do while(i_ret==CELLTYPE_NUMBER)
         iR=iR+1; i_ret=xlSheetCellType(hArk,iR,0)
         if(i_ret==CELLTYPE_STRING.or.i_ret==CELLTYPE_EMPTY.or.i_ret>=CELLTYPE_BLANK)then; 
            Rend=num2str(iR);       !1-based Number = iR-th Row#
            exit;                   !0-based Row Number = iR which is a string cell =>'Mean'
         end if
        end do !CELLTYPE_EMPTY=0,CELLTYPE_NUMBER=1,CELLTYPE_STRING=2,CELLTYPE_BOOLEAN=3,CELLTYPE_BLANK=4,CELLTYPE_ERROR=5
        nTot=iR-iR0; call reallocate(ColData,nTot,nCol); call reallocate(DumArr,nTot)
        do i=iR0,iR-1
           DumArr(i-iR0+1)=xlSheetReadNum(hArk,i,0)! for 0-based indices           
           do II=1,nCol
            ColData(i-iR0+1,II)=xlSheetReadNum(hArk,i,Col_result(II))
           end do
        end do;
        Twind=DumArr(nTot)-DumArr(1)-TransT; dt=DumArr(2)-DumArr(1); nTr=int(TransT/dt)+1;
If(nHr>0)then
  Twind=nHr*3600.
  n3Hr=int(DumArr(nTot)/Twind)
  nT_3h=int(Twind/dt)
end if
      ! Note: DumArr will be reallocated to have smaller size starting with k0:k1
       DO k=1,n3Hr
          k0=nTr+nT_3h*(k-1); k1=min(k0+nT_3h-1,nTot)
          do j=1,nCol
           if(maxval(ColData(k0:k1,j))<0)then; iSign=-1; else; iSign=1; endif; yTitle=ReadCellStr(hArk,iR0-1,j)
           if(iMaxType>0)then
             call WeiBulTailExtreme(iSign*ColData(k0:k1,j),Twind,xExt,pValue,WeibPara)!,yTitle
             MeanStdMinMax=[mean(ColData(k0:k1,j)),std(ColData(k0:k1,j)),minval(ColData(k0:k1,j)),iSign*xExt(1)]
           else  
             MeanStdMinMax=[mean(ColData(k0:k1,j)),std(ColData(k0:k1,j)),minval(ColData(k0:k1,j)),iSign*maxval(iSign*ColData(k0:k1,j))]
           end if  
           forall(i=1:4)Table(k,(nLin+size(Col_Noddis)+2+j)*4+i)=MeanStdMinMax(i)
          end do
       END DO
        ErrMsg=GetErrorMsg(hBuk); call xlBookRelease(hBuk)
        if(ErrMsg/='ok')call ErrorMsg('Error in reading sheet:'//XLFileList(iXL),5)
      !END if
     write(*,'(a)')'Result Data Processed! Now writing to : '//XLfil//'->Sheet:'//ArkNam0//' & '//ArkNam
      I=len_trim(VarList(iLS,nL))-2; J=len_trim(path)-1; ErrMsg=VarList(iLS,nL)(:I)//', MaxType='//merge('WTF','Obs',iMaxType>0)
      Shnam=Path(scan(path(:j),'\',.true.)+1:j)//ErrMsg; tmpstr='Dummy'; j=0
      call xlSheetWriteComment(Arkhdl0,0,j,Path//Shnam//""C)
      do while(len_trim(tmpstr)>0)
       tmpstr=ReadCell(Arkhdl0,0,j);
        if(instr(tmpstr,Shnam))then
          call WriteCell(Arkhdl0,0,j,Shnam);
          call WriteCell(Arkhdl,0,j,Shnam);
        end if; j=j+1
      end do
      fName=' '
      DO k=1,n3Hr
       if(n3Hr>1)then
         fName=OrderNum(k)//'_3Hr'  
       end if    
       call WriteCell(Arkhdl0,iRow,0,trim(VarList(iLS,1))//fName);
       call WriteCell(Arkhdl,iRow,0, trim(VarList(iLS,1))//fName);
       call WriteCell(Arkhdl0,iRow,1,Table(k,:));
       LTmx=LTen(k,:,2); j1=maxloc(LTmx(:)); LTmx(j1)=0d0; j2=maxloc(LTmx(:)); 
       call WriteCell(Arkhdl,iRow,1,[LTen(k,j1,:),MBL/LTen(k,j1,2),dble(j1),LTen(k,j2,1:2),dble(j2),(Table(k,4*nLine+1+4*(i-1):4*nLine+1+4*i:3),i=1,3)]);
       iRow=iRow+1
     END DO
     i_ret=xlBookSave(Bukhdl, XLfil//""C); ErrMsg=GetErrorMsg(Bukhdl);
     if(ErrMsg=='ok')write(*,'(a)')'Finished writing in '//XLfil
    !END IF
  END DO

End subroutine SIMAXLSData

 subroutine CommandLine()
  use nrtype,only:CR
 use help_mod
      !interface
      !  subroutine help(Title)
      !   character(*),intent(in),optional::Title
      !  end subroutine help
      !end interface
        
        integer::argc,status,NEXT,res,lret,iUfil,i,j
        character*50::argv
        character(:),allocatable::crad,LogFil,SumFil,str0,str1,strMSG(:)
        integer::Arglen,nch,ios,iDebug,nArg
        logical(4)::bMoreParams     ! flag end of new process parameter processing
        bMoreParams = .TRUE.; Arglen=1; iDebug=5; TransT=3.*60.*sqrt(50.)! Transient time about 3 min in model scale
        pValue = exp(-1d0)
        NEXT=1; nLC=0; MBL=16500; nLine=12; iMaxType=0; nHr=0; nArg=command_argument_count()
        !SimaPp.exe /nNoodis.txt /eElemfor.txt /rResults.txt /wWiturb.txt /lNoddis_log.txt /lWiturb_log.txt

        do while(bMoreParams .NEQV. .FALSE.)
            crad=' '; CALL get_command_argument(NEXT,crad,nch,ios); ios=-ios;
            if(ios/=0)then;
                crad=repeat(' ',nch); call getarg(NEXT,crad,IOS); ! argv++ ; point to the first parameter
            else
                ios=nch
            endif;
            ! lret=MessageBox(ghwndmain,crad//"|ios="//num2str(ios)//""C,OrderNum(NEXT)//"/"//num2str(command_argument_count())//" Arg"C,MB_OK)
            !if(NEXT>10)write(*,*)'Arg#',NEXT,'is ',crad, ', ios=',ios
            NEXT=NEXT+1;
            !if(ios/=#FFFF)crad(max(ios+1,1):max(ios+1,1))=char(0)
            if(ios==-1)then ! if*crad is NULL we're missing the program name!
                !------------------------- debug... --------------------------
#IF DEFINED(_DEBUG)
                 !if(NEXT==2)crad='/p.\'
                 !if(NEXT==2)crad='/pD:\WorkDir\ProjSpace\M2269_B29\model_test\data\irrwave_timeseries'!'/pD:\WorkDir\ProjSpace\M2269_B29\Sima46\MTL_QTF\CondSet\37-20241108130142'
                 !if(NEXT==)crad='/l3100'
                 !if(NEXT==)crad='/l3120'
                 !if(NEXT==)crad='/l3170'
                 !if(NEXT==)crad='/l3180'
                 !if(NEXT==3)crad='/l3190'
                 !if(NEXT==4)crad='/sW'
                 !if(NEXT==5)crad='/c1390.'
                 !if(NEXT==6)crad='/xOSXMTemplate.xlsx';
!!======================== SIMA_Results ......................................................................
if(NEXT==2)crad='/pD:\WorkDir\ProjSpace\M2114_Wisting\Sima\SRE_LowMLBE_Ballast\B_ULS_Sd\LC10x3h\rLC2_dt16\node_1'
if(NEXT==2)crad='/pD:\WorkDir\ProjSpace\2408SSY_ARFLNG\Sima_HBEModel\ArgFLNG_v0_35m_Loaded\condset\240-20250813081948\condset_1'
if(NEXT==3)crad='/lCondset1'
if(NEXT==4)crad='/xD:\WorkDir\ProjSpace\M2114_Wisting\Sima\SRE_LowMLBE_Ballast\B_ULS_Sd\LC10x3h.xlsx'
if(NEXT==4)crad='/x"D:\WorkDir\ProjSpace\2408SSY_ARFLNG\Sima_HBEModel\ArgFLNG_v0_35m_Loaded\condset\240-20250813081948\condset_1\key_sima_elmfor.txt.xlsx"'
if(NEXT==5)crad='/tSIMA SIMAPlt'
if(NEXT==6)crad='/sX'
if(NEXT==7)crad='/c1000.';
if(NEXT==8)crad='/MObs';
if(NEXT==9)crad='/MBL=19748';
if(NEXT==10)crad='-n12';
if(NEXT==11)crad='/M3hr';
!if(NEXT==2)crad='/pD:\WorkDir\ProjSpace\M2114_Wisting\Sima\Simulations_SRE\L_ALS1F_Sd1yr\rLC1'
!if(NEXT==3)crad='/lALS1F_1yr_LC1'
!if(NEXT==4)crad='/xModelTestResSum.xlsx'
!if(NEXT==5)crad='/tSIMA46 SIMA46Plt'
!if(NEXT==6)crad='/sX'
!if(NEXT==7)crad='/c1390.';
!if(NEXT==8)crad='/MObs';
!if(NEXT==9)crad='/MBL=19748';

                 !if(NEXT==10)crad='/tSIMA_Data NumExtremeSum'
                 !if(NEXT==11)crad='/sX';
                 !if(NEXT==12)crad='/MBL=16500';
                 !if(NEXT==13)crad='/xModelTestResSum.xlsx';
             !if(NEXT==3 )crad='/l3011 3120'
             !if(NEXT==4 )crad='/l3020 3130'
             !if(NEXT==5 )crad='/l3030 3140'
             !if(NEXT==6 )crad='/l3041 3150'
             !if(NEXT==7 )crad='/l3050 3160'
             !if(NEXT==8 )crad='/l3060 3170'
             !if(NEXT==9 )crad='/l3070 3180'
             !if(NEXT==14)crad='/sw';
                 !if(NEXT==15)crad='/c1390.';
! if(NEXT==2)crad='/p.\'
! if(NEXT==3 )crad='/l3041 3060'
! if(NEXT==4 )crad='/l3080 3070'
! if(NEXT==5 )crad='/l3110 3090'
! if(NEXT==6 )crad='/l3140 3100'
! if(NEXT==7 )crad='/l3201 3120'
! if(NEXT==8 )crad='/l3210 3130'
! if(NEXT==9 )crad='/l3220 3150'
! if(NEXT==10)crad='/l3011 3160'
! if(NEXT==11)crad='/l3020 3170'
! if(NEXT==12)crad='/l3030 3180'
! if(NEXT==13)crad='/l3050 3190'
! if(NEXT==14)crad='/sI';!'/sw';!
! if(NEXT==15)crad='/MObs';!'/MWFT'!
! if(NEXT==16)crad='/c1390.'
! if(NEXT==17)crad='/xModelTestResSum.xlsx';
! if(NEXT==18)crad='/tCollectMObs MTMObsSum'
 !if(NEXT==18)crad='/tCollection MTExtremeSum'
#ELSE
                exit;  call ExitProcess(1)
#ENDIF                
            end if;
            nch=clen(crad)
            if(nCh<3)then
              if(nCh==0)exit
              cycle
            else
                strMSG=[strMSG,'Argument Captured #'//OrderNum(NEXT-1)//"/"//num2str(nArg)//":"//crad]
            end if    
            if((crad(1:1) == '/') .OR. (crad(1:1) == '-')) then
                select case (crad(2:2))!((*crad)[1])  letter after the '/' or '-'
                case ('p')  ! Path
                  Path=crad(3:); if(crad(nch:nch)/='\')Path=Path//'\'
                case ('x')  ! Excel output file
                  XLfil=crad(3:)
                case ('l')  ! Label input
                  if(crad(3:3)=='"'.and.crad(nch:nch)=='"')then
                    crad=crad(:2)//crad(4:nch-1); nch=nch-2
                  end if
                  i=scan(Crad(3:),' ',.true.)+3; 
                  if(crad(nch:nch)/='\')then
                    str1=crad(i:nch)//'\'
                  else   
                    str1=crad(i:nch);
                  end if 
                  if(i>4)then
                     str0=crad(3:i-2);
                     StrArr=[StrArr,str0,str1];
                  else  
                     StrArr=[StrArr,str1];
                  end if  
                case ('c','T')! Cut time
                    read(crad(3:nch),*,iostat=ios)TransT
                    if (ios /= 0) call help("Invalid /c (cuttime) value")
                case ('n')  ! Number of lines
                    read(crad(3:),*,iostat=ios)nLine
                    if (ios /= 0) call help("Invalid /n (mooring lines) value")
                case ('M')  ! Max type, M3hr or MBL=...
                    i=index(crad(3:),'hour')+2; 
                    if(i==2)i=index(crad(3:),'hr')+2;
                    if(i>2)then!/M3hr !3 hour maximum used to truncate the long simulation into several 3h maxima.
                     read(crad(3:i-1),*,iostat=ios)nHr
                     if(ios/=0)call help("Invalid value for number of 3Hr count, use /M3hr or /M3Hour!")
                    end if
                    if(instr('MP',crad(2:4)))then
                      if(index(crad,'MPM')>1 .or. index(crad,'MP37')>1 .or. index(crad,'MP36.7')>1 )then
                        pValue = exp(-1d0)
                      elseif(index(crad,'MP57')>1 .or. index(crad,'MEM')>1 .or.  index(crad,'MELM')>1 .or. index(crad,'MExpectedMaximum')>1 )then
                        pValue = 0.5703760016750230369575524240486990574628d0 !ELMP=e^(-e^EULER)
                      elseif(scan(crad(3:),'0123456789')>0)then
                        i=scan(crad(4:),'0123456789')+3
                        read(crad(index(crad,'=')+1:),*,iostat=ios)pValue
                        if(ios/=0)read(crad(i:),*,iostat=ios)pValue
                        if(ios/=0)call help("Invalid value for extreme value Percentile, use MPM/MP37, MELM/MEM/MP57!")
                      end if
                    end if
                    if(nch>=5)then
                        i=scan(crad(5:),'0123456789')+4; ios=-1
                        if(i>4)then
                          if(INDEX(crad, 'MBS') > 1 .or. INDEX(crad, 'MBL') >1 .or. INDEX(crad, 'MinimumBreakingLoads') >1 .or. &
                             INDEX(crad, 'MinimumBreakingStrength') >1 .or. INDEX(crad, 'MinBreakingStrength') >1)&
                             read(crad(index(crad,'=')+1:),*,iostat=ios)MBL
                             if(ios/=0)read(crad(i:),*,iostat=ios)MBL
                             if(ios/=0)call help("Invalid value for MBS/MBL/MinimumBreakingLoads/MinimumBreakingStrength/MinBreakingStrength!")
                        elseif(instr('MObs,MaxObs,MObserved',crad(2:4)))then! Take Observed Maximum
                         iMaxType=0    
                        elseif(instr('MWTF,MWFT,MaxWeibull,MWeiBullFIT,MWeiBullTailFIT',crad(2:4)))then! Take WeibulFitted Maximum
                         iMaxType=1
                        end if    
                    end if
                    !if(nch<6.or.ios/=0)call help("Error: invalid switch for MBL of Lines: /"//crad)
                case ('t')  ! Sheet names (expect next argument too)
                  i=scan(Crad(3:),' ',.true.)+3
                  if(i>4)then
                   ArkNam0=crad(3:i-2); ArkNam=crad(i:nch); 
		          else 
                   ArkNam0 = crad(3:)
                   call get_command_argument(NEXT, crad, nch, ios)
                   if (ios == 0) then
                        ArkNam = adjustl(crad(1:nch))
                        NEXT = NEXT + 1
                   else
                        call help("Missing second sheet name for /t")
                   end if
                  end if
                case ('s') ! Series type switch! Use the time series from Sintef Ocean Model test excel data (with ObsMaximum) to WeibulMaximum 
                    select case (crad(3:3))
                     case('w')! WeiBullFitXLData(XLfil,ArkNam//''C,TransT): Insert WeibullFitted max to Timeseries xls file. 
                        iType=-1; ! 
!ModelTestResSum /l3011 /l3020 /l3030 /l3041 /l3050 /l3060 /l3070 /l3080 /l3090 /l3110 /l3130 /l3140 /l3150 /l3160 /l3201 /l3210 /l3220 /sw /c1390.
                     case('W')! ,'WF','Wf','wf','WeibulFit')
                        iType=1; !call IrrWaveTestSummary
                        nLC=size(StrArr)
                        do i=1,nLC
                          j=clen(StrArr(i))
                          if(StrArr(i)(j:j)=='\')StrArr(i)=StrArr(i)(:j-1) !test%i.sta.xlsx
                        end do
!for %i in (3011 3020 3030 3050 3060 3070 3090 3100 3120 3130 3041 3080 3110 3140 3150 3160 3170 3180 3190 3201 3210 3220)do (copy/y T0.sta.xlsx test%i.sta.xlsx&Binfile test%i.sta -1 test%i.sta.xlsx)
                     case('X','x')
                        iType=2
!if(iType==2)call SIMAXLSData(VarList,Path,XLfil,ArkNam,Arkhdl0,Arkhdl)
!set p=D:\WorkDir\ProjSpace\M2269_B29\Sima46\MTB_QTF_Simo426\CondSet\24-20241111100837
!ModelTest2Xls "/p%p%" "/l3041Wd[.35,.4] CondSet_1" "/l3080Wd[.35,.4] CondSet_2" "/l3110Wd[.35,.4] CondSet_3" "/l3140Wd[.35,.4] CondSet_4" "/l3201Wd[.35,.4] CondSet_5" "/l3210Wd[.35,.4] CondSet_6" /xModelTestResSum.xlsx "/tSIMA_Data NumExtremeSum" /sX /MBL=16500
!set p=D:\WorkDir\ProjSpace\M2269_B29\Sima46\MTB_Newman\CondSet\33-20241113082624
!ModelTest2Xls "/p%p%" "/l3041Wd[.3,.4] CondSet_1" "/l3080Wd[.3,.4] CondSet_2" "/l3110Wd[.3,.4] CondSet_3" "/l3140Wd[.3,.4] CondSet_4" "/l3201Wd[.3,.4] CondSet_5" "/l3210Wd[.3,.4] CondSet_6" /xModelTestResSum.xlsx "/tSIMA46 SIMA46Plt" /sX /MBL=16500
                     case('I','i')
                        iType=3
!if(iType==3)call SIOModelTestXLSData(VarList,Path,XLfil,ArkNam,Arkhdl0,Arkhdl)
!Example run:
!set p=..\ModelTest2Xls\
!ModelTest2Xls "/p%p%" "/l3041 3020" "/l3080 3030" "/l3110 3050" "/l3140 3060" "/l3201 3070" "/l3210 3090" "/l3011 3201" "/l3080 3210" "/l3100 3220" "/l3110 3190" /sI xModelTestResSum.xlsx "/tCollection MTExtremeSum"
                     case('M','m')
                        iType=4
!if(iType==4)call MarinModelTestCSVData(VarList)
                     case DEFAULT
                      
                    end select
                case ('?','h')
                    call help() !  help() will terminate app
                case DEFAULT
                 call help("Error: invalid switch - "//crad)
                end select
                !           strMSG=trim(TabT1)//'|2'//trim(TabT2)//'|3'//trim(TabT3)//'|4'//trim(TabT4)
                !           PBres=ProgDialog(0,"SimaPp",.true.,strMSG,IDC_OKOK,"Continue?"C,.true.,10); if(PBres==2)Stop

            else ! not a '-' or '/', must be the program name
                bMoreParams = .FALSE.
            end if
        end do
        write(*,*)strMSG
        end subroutine CommandLine


End Program ModelTest2XL