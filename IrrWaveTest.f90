subroutine IrrWaveTestSummary(Dir,iCol2Cal,iCo2Merg,XLfil,SumXLFil,TransT)
USE IFPORT; USE,INTRINSIC::ISO_FORTRAN_ENV,ONLY:IOSTAT_EOR
use kernel32,only:MoveFileEx
use util_xk,only:reallocate,swap,mean,std
use LStrFileUti
use libxl
use StatPack,only:WeiBulTailExtreme!,LevelCrossing,StatPara,WeibullMLE4GB,WeibullMME_Skewness4B,FITDIS
IMPLICIT  NONE
character(*),intent(in)::Dir,XLfil,SumXLFil
integer,allocatable::iCol2Cal(:,:),iCo2Merg(:,:)
real,intent(in)::TransT
type(BookHandle)::Bukhdl0,BukHdl
type(SheetHandle)::Arkhdl0,Ark0h,Arkhdl
type(FormatHandle)::Centerfmt,newFmt,fmt1,LTRB1fmt,LTRB2fmt,NumD2Fmt,NumD2FmtBkgClr,Fmt(4)
type(FontHandle)::font1
integer::nfmt,nfont,iBukLoad,iBukOld,i_ret,rnum,cnum,itype
character(:),allocatable::crad,filnam,ArkNam0,ArkNam,TabHdr(:),ColName,tempFile,colstr
integer,parameter::nLC=1,nLine=9,MBL=10200
integer::iui=1,iflog=2,iuo=3,i,ir,II,j,jc,k,m,ios,size,nRec,nStep,nArg,iRow,iCol,nCol,nStr,nNum,nCol2Use,iMea,iStd,iMin,iMax,j1(1),j2(1),nCh2Asc
integer,allocatable::ipos(:),iColmn(:),iNewOrd(:)
CHARACTER(256)::str*20,arg,strL*500,str2,iomsg,n2s*2,DateStr*9
character(:),allocatable::path,iFil,oFil,VarList(:,:),FILOUT(:),StrArr(:),VarHdr(:),ErrMsg,Rend,Rbeg
real*8,allocatable::NumArr(:),Table(:,:),LTen(:,:),LTmx(:)
! Stats
integer,parameter::nMom=4,nCLS=10; 
real*8,DIMENSION(:),allocatable::X,XParent,XMaxima,XMinima,Pr,pdf,CDF
real*8,DIMENSION(:,:),allocatable::xEXt
real*8::LOC_OLD,SCA_OLD,SHA_OLD,PLKey,P_ext,rMax,rMin,rMean,xMean,Twind,rStd,rSkew,rKurt,CLsz,gama,beta,mu,xMomc(0:nMom),GaBeMu(3),MLEpar(1),MMEPar(2),StatPar(6)
character(:),allocatable::fName,xTitle,yTitle
integer::pUnique,ErrType,iWei,nMxNo,nTr
!pValue: Statistic percentile=>any one of the hundred groups so divided while fractile is the value of a distribution for which some fraction of the sample lies below.
real*8::pValue=.37d0,pLow=.80d0,MxTmp!=PFRAC; else; pValue=.37d0
!if(pres_PLKey)then; pLow=PLKey; else; pLow=.80d0; endif

nCol2Use=38; tempFile='tmp_1'C; fname='WeibullTailplot.plt'; xTitle='Line Tension X';yTitle='-ln(1-F)'
i=MakeDir('WFIT_plt');
Bukhdl0=NEWXMLBook(); 
iBukLoad=xlBookLoad(Bukhdl0,ctrim(Dir)//ctrim(XLfil)//''C);
!iBukload=xlBookLoadPartiallyUsingTempFile(Bukhdl0,ctrim(Dir)//ctrim(XLfil)//""C,0,0,125904,tempFile,1)
IF(iBukLoad==1)then; i=index(XLFil,'.sta'); if(i==0)i=len(XLFil); j=index(XLFil,'.xls'); if(j==0)j=len(XLFil)
   k=index(XLFil,'test')+4; m=min(i,j)-1
   if(k==4)k=1
   ArkNam=XLFil(k:m)
   Arkhdl0=xlBookGetSheetbyName(Bukhdl0,'0'//ArkNam//''C)
   if(Arkhdl0%Point==0)Arkhdl0=xlBookGetSheetbyName(Bukhdl0,ArkNam//''C)
   if(Arkhdl0%Point==0)then
       call ErrorMsg('Error in loading sheet:'//ArkNam//' of '//XLfil,-1)
   else
    Bukhdl=NEWXMLBook(); 
    iBukOld=xlBookLoad(Bukhdl,ctrim(SumXLfil)//""C)
    !iBukOld=xlBookLoadPartiallyUsingTempFile(Bukhdl,ctrim(SumXLfil)//""C,0,0,125904,tempFile,1)
    !Bukhdl=NEWXLBook(); iBukOld=xlBookLoad(Bukhdl,ctrim(SumXLfil)//""C) !OK to combine .xls files
    !!NumD2fmt=xlBookAddFormat(Bukhdl,nullformat);   call xlFormatSetNumFormat(NumD2fmt,NUMFORMAT_NUMBER_D2)
    !!NumD2fmtBkgClr=xlBookAddFormat(Bukhdl,NumD2fmt); 
    !CALL xlFormatSetFillPattern(NumD2fmtBkgClr,FILLPATTERN_GRAY50)
    !CALL xlFormatSetPatternForegroundColor(NumD2fmtBkgClr,COLOR_TAN)
    if(iBukOld==0)then
      ErrMsg=GetErrorMsg(Bukhdl);
      call ErrorMsg(ErrMsg//' in loading Existing summary file:'//SumXLfil,-1)
    else
 ! ========== This is temporarily copy Weibull fitted MaxT in the test0000.sta.xlsx file to test0000.xlsx
      !Arkhdl=xlBookGetSheetbyName(Bukhdl,ArkNam//""C); 
      !iR=100000;jc=0; call FindCell(Arkhdl0,['MaxT'],iR,jc)
      !  do i=iRow,iRow+1
      !     do j=jc+1,jc+nCol2Use
      !      MxTmp=xlSheetReadNum(Arkhdl0,i,j);
      !      call WriteCell(Arkhdl,i,j,MxTmp)
      !     end do
      !  end do
      !  i_ret=xlBookSave(Bukhdl,ArkNam//".xlsx"C)
      !  call xlBookRelease(Bukhdl0); 
      !  call xlBookRelease(Bukhdl); 
      !stop    
 ! ========== This is temporarily copy Weibull fitted MaxT in the test0000.sta.xlsx file to test0000.xlsx
!   Add Header & Footer to each LoadCase file together with WEIBUL fit Extreme
      Ark0h=xlBookGetSheetbyName(Bukhdl,'0'); call xlSheetSetName(Ark0h,ArkNam//''C)
      if(Ark0h%Point==0)then
         call ErrorMsg('Error in loading Template sheet: "0" of '//SumXLfil,-1)
      else
        Arkhdl=xlBookGetSheetbyName(Bukhdl,ArkNam//""C);
        if(Arkhdl%Point==0)Arkhdl=xlBookAddSheet(Bukhdl,ArkNam//''C)
      end if
      if(Arkhdl%Point==0)then
          call ErrorMsg('Error in adding sheet:'//ArkNam//' of '//SumXLfil,-1)
      else
!       Now the Dest.Excel is ready, just need to read the Source data from ArkHdl0
        iRow=3; i=iRow; i_ret=xlSheetCellType(Arkhdl0,i,0);
        do while(i_ret==CELLTYPE_NUMBER)
         i=i+1; i_ret=xlSheetCellType(Arkhdl0,i,0)
         if(i_ret==CELLTYPE_STRING.or.i_ret==CELLTYPE_EMPTY.or.i_ret>=CELLTYPE_BLANK)then; 
             Rend=num2str(i+3);     !1-based Number = ith Row in Arkhdl0 (3041.sta.xlsx) (i+3)th Row in 3041.xlsx (+3 Header Rows)
             iRow=i-1; exit;        !0-based Row Number = i - 1
         end if
        end do !CELLTYPE_EMPTY=0,CELLTYPE_NUMBER=1,CELLTYPE_STRING=2,CELLTYPE_BOOLEAN=3,CELLTYPE_BLANK=4,CELLTYPE_ERROR=5
        IF(size(iCol2Cal)==0)then
         j=1; i_ret=xlSheetCellType(Arkhdl0,iRow,j);
         do while(i_ret==CELLTYPE_NUMBER)
          j=j+1; i_ret=xlSheetCellType(Arkhdl0,iRow,j)
          if(i_ret==CELLTYPE_STRING.or.i_ret==CELLTYPE_EMPTY.or.i_ret>=CELLTYPE_BLANK)then; 
              nCol2Use=j-1; exit;        !0-based Number = j - 1
          end if
         end do !CELLTYPE_EMPTY=0,CELLTYPE_NUMBER=1,CELLTYPE_STRING=2,CELLTYPE_BOOLEAN=3,CELLTYPE_BLANK=4,CELLTYPE_ERROR=5
        ELSE
          nCol2Use=sum(iCol2Cal(:,2)-iCol2Cal(:,1)+1)  
        END if
        if(allocated(Table))deallocate(Table); Allocate(Table(iRow-2,nCol2Use),stat=ios)!call reallocate(Table,iRow-2,nCol2Use)
        i=1; ColStr=''
        do while(ios>0)
          Allocate(Table(iRow-2,nCol2Use),stat=ios)
          call ErrorMsg('Error in Memory allocation: Alloc_ERR='//num2str(ios)//'!Retry#'//num2str(i),2*i)
        end do
        if(allocated(xExt))deallocate(xExt); Allocate(xExt(2,nCol2Use-1))
        DO ir=3,iRow
          j=0  
          do II=1,size(iCol2Cal,1)
           do jc=iCol2Cal(II,1),iCol2Cal(II,2)
            j=j+1  
            Table(ir-2,j)=xlSheetReadNum(Arkhdl0,ir,jc);
           end do
          end do
        END DO; Twind=Table(iRow-2,1)-Table(1,1)-TransT; nTr=int(TransT/(Table(2,1)-Table(1,1)))+1
        call WriteCell(Arkhdl,6,0,Table); ; Rbeg=num2str(nTr)
        do i=0,5
         crad=ReadCellStr(Ark0h,i,0,fmt1)
         call WriteCell(Arkhdl,i,0,crad,fmt1)
         if(i<4)call WriteCell(Arkhdl,i+iRow+4,0,crad,fmt1)!Row for 'Max' is iRow+7
        end do
        call WriteCell(Arkhdl,3+iRow+5,0,'MaxT',fmt1)
        call WriteCell(Arkhdl,3+iRow+6,0,'-LOG(1-P_ext)',fmt1)
        do i=4,5
         do jc=1,nCol2Use
         crad=ReadCellStr(Ark0h,i,jc,fmt1)
         call WriteCell(Arkhdl,i,jc,crad,fmt1)
         end do
        end do
      do jc=1,nCol2Use-1!0-based Column Number = jc=1 for 2nd Column (B)
      !do jc=28,nCol2Use-1!0-based Column Number = jc=1 for 2nd Column (B)
        ColName=trim(XLCol2Str(jc+1))
        call WriteCell(Arkhdl,0,jc,mean(Table(nTr:,jc+1)));
        call WriteCell(Arkhdl,1,jc,std(Table(nTr:,jc+1)));    
        call WriteCell(Arkhdl,2,jc,minval(Table(nTr:,jc+1)));
        call WriteCell(Arkhdl,3,jc,maxval(Table(nTr:,jc+1)));
 !     i_ret=xlSheetReadFormula(Arkhdl,0,jc,fmt(1)); i_ret=xlSheetWriteFormula(Arkhdl,0,jc,"=AVERAGE("//ColName//'7:'//ColName//Rend//")"C,fmt(1));
 !     i_ret=xlSheetReadFormula(Arkhdl,1,jc,fmt(2)); i_ret=xlSheetWriteFormula(Arkhdl,1,jc,"=STDEV("//ColName//'7:'//ColName//Rend//")"C,  fmt(2));
 !     i_ret=xlSheetReadFormula(Arkhdl,2,jc,fmt(3)); i_ret=xlSheetWriteFormula(Arkhdl,2,jc,"=MIN("//ColName//'7:'//ColName//Rend//")"C,    fmt(3));
 !     i_ret=xlSheetReadFormula(Arkhdl,3,jc,fmt(4)); i_ret=xlSheetWriteFormula(Arkhdl,3,jc,"=MAX("//ColName//'7:'//ColName//Rend//")"C,    fmt(4));
        i_ret=xlSheetWriteFormula(Arkhdl,0+iRow+4,jc,"=AVERAGE("//ColName//Rbeg//':'//ColName//Rend//")"C,fmt(1));
        i_ret=xlSheetWriteFormula(Arkhdl,1+iRow+4,jc,"=STDEV("//ColName//Rbeg//':'//ColName//Rend//")"C,  fmt(2));
        i_ret=xlSheetWriteFormula(Arkhdl,2+iRow+4,jc,"=MIN("//ColName//Rbeg//':'//ColName//Rend//")"C,    fmt(3));
        i_ret=xlSheetWriteFormula(Arkhdl,3+iRow+4,jc,"=MAX("//ColName//Rbeg//':'//ColName//Rend//")"C,    fmt(4));
!        X=Table(:,jc+1);
!! The following is to build up the statistic probility distribution model for WeiBull fitting
!!     For Long-term 100yr ULS Extreme: F(x)^N=1-1/100=>F(x)=(1-.01)^(1/N); Annual probability of non-exceedance for N (independent events per year), e.g.N=20/35(20 storms in 35 years)
!!     For short-term 3-hr SS Extrem:F(x)^N=PFRAC=>F(x)=PFRAC^(1/N);Non-exceedance PFRAC for N peaks and no# of 3hr,=PFRAC**(TimeWinUpp-TimeWinLow)/(NoPeaks*PERIOD)          
!!     The exceedance probability 1/N of a single 3-hour maximum to exceed the value for the limit states are: 1/(100*2922) and 1/(10000*2922) for 100yr-ULS and 10000yr-ALS respectively.
!        call StatPara(nCLS,X,rMin,rMean,rMax,rStd,rSkew,rKurt,'ParentSeries.txt')
!        call LevelCrossing(X,XMaxima); nMxNo=size(XMaxima); Pr=[(dble(i)/(nMxNo+1d0),i=1,nMxNo)]; P_ext=pValue**(Twind/(nMxNo*ThreeHr))
!        call StatPara(nCLS,xMaxima,rMin,rMean,rMax,rStd,rSkew,rKurt,'MaxSeries.txt'); LOC_OLD=rMin-1d0/nMxNo; SCA_OLD=225.59162d0; SHA_OLD=1.53248d0
!!     call reallocate(xMaxima,nPnt); xMaxima=X; Pr=XX(ir+2,:); P_ext=.99999d0
!        crad=ReadCellStr(Arkhdl,4,jc); oFil=ReadCellStr(Arkhdl,5,jc)
!        if(clen(crad)>0)ColStr=ctrim(crad); xTitle='WFIT_plt\'//ArkNam//ColStr//oFil
!        call FITDIS(xMaxima,Pr,P_ext,pLow,pUnique,ErrType,iWei,xExt(:,jc),LOC_OLD,SCA_OLD,SHA_OLD,xTitle//fName,xTitle,yTitle); 
!        close(114); i=MoveFileEx('fort.114',xTitle//'FitResp.txt'C,0)
        call WeiBulTailExtreme(Table(nTr:,jc+1),Twind,xExt(:,jc))
        call WriteCell(Arkhdl,3+iRow+5,jc,xExt(:,jc),fmt(4),'Row')
      end do
      if(allocated(iPos))deallocate(iPos); allocate(iPos(4))
      do II=1,size(iCo2Merg,1)
       if(xlSheetGetMerge(Arkhdl,4,iCo2Merg(II,1),iPos(1),iPos(2),iPos(3),iPos(4))==0)then!GetMerge=1 if specified cell is in a merged area else returns 0
         i_ret=xlSheetSetMerge(Arkhdl,4,4,iCo2Merg(II,1),iCo2Merg(II,2))!For existing excel file, MUST use xlSheetDelMerge(handle,row,col) first!
       end if
      end do
      i_ret=xlBookSave(Bukhdl,ArkNam//".xlsx"C)
     end if 
    end if
    call xlBookRelease(Bukhdl); 
   end if 
ELSE
   ErrMsg=GetErrorMsg(Bukhdl0);
   call ErrorMsg(ErrMsg//' in loading Existing summary file:'//ctrim(Dir)//ctrim(XLfil),-1)
END IF
call xlBookRelease(Bukhdl0); 
End subroutine IrrWaveTestSummary
    
subroutine IrrWaveTest
USE IFPORT; USE,INTRINSIC::ISO_FORTRAN_ENV,ONLY:IOSTAT_EOR
use util_xk,only:reallocate,swap,mean,std
use LStrFileUti
use libxl
IMPLICIT  NONE
type(BookHandle)::Bukhdl
type(SheetHandle)::Arkhdl0,Arkhdl
type(FormatHandle)::Centerfmt,newFmt,fmt1,LTRB1fmt,LTRB2fmt
type(FontHandle)::font1
integer nfmt,nfont,iBukLoad,i_ret,itype
character(:),allocatable::crad,filnam,XLfil,ArkNam0,ArkNam,TabHdr(:)
integer,parameter::nLC=1,nLine=9,MBL=10200
integer::iui=1,iflog=2,iuo=3,i,II,j,k,m,ios,nRec,nStep,nArg,iRow,iCol,nCol,nStr,nNum,nCol2Use,iMea,iStd,iMin,iMax,j1(1),j2(1),nCh2Asc
integer,allocatable::ipos(:),iColmn(:),iNewOrd(:)
CHARACTER(256)::str*20,arg,strL*500,str2,iomsg,n2s*2,DateStr*9
character(:),allocatable::path,iFil,oFil,VarList(:,:),FILOUT(:),StrArr(:),VarHdr(:)
real*8,allocatable::NumArr(:),Table(:,:),LTen(:,:),LTmx(:)
!allocate(character(64)::tmpStr)
!VarHdr=['TIME','F_MOOR.01_P','F_MOOR.02_P','F_MOOR.03_P','F_MOOR.04_P','F_MOOR.05_P','F_MOOR.06_P','F_MOOR.07_P','F_MOOR.08_P','F_MOOR.09_P','FX_TUR_C','FY_TUR_C','FXY_TUR_C','FZ_TUR_C','X_TUR_SWIV','Y_TUR_SWIV','XY_TUR','Z_TUR_SWIV','X_COG','Y_COG','Z_COG','ROLL','PITCH','YAW','WAVE_180','WAVE_210','WAVE_270','CUR_REF','F_PR.01_P','F_PR.02_P','F_PR.03_P','F_PR4.04_P','F_GER.01_P','F_GER.02_P','F_GER.03_P','R','V','Hx','Hy','Hxy']
!====== Turret Force Comparison vs Model Test Results show Inertia Corrected becomes larger than Numerical ========
!======  Turret Forces are taken as original without Modification. 24.02.2020
VarHdr=['TIME','F_MOOR.01_P','F_MOOR.02_P','F_MOOR.03_P','F_MOOR.04_P','F_MOOR.05_P','F_MOOR.06_P','F_MOOR.07_P','F_MOOR.08_P','F_MOOR.09_P','FX_TUR','FY_TUR','FZ_TUR','X_TUR_SWIV','Y_TUR_SWIV','XY_TUR','Z_TUR_SWIV','X_COG','Y_COG','Z_COG','ROLL','PITCH','YAW','WAVE_180','WAVE_210','WAVE_270','CUR_REF','FX_TUR_C','FY_TUR_C','FXY_TUR_C','FZ_TUR_C','F_PR.01_P','F_PR.02_P','F_PR.03_P','F_PR4.04_P','F_GER.01_P','F_GER.02_P','F_GER.03_P','R','V','Hx','Hy','Hxy']
Path="X:\analysis\projects\E2128 Reliance\ModelTest\ftp.Marin.nl\4_Results\Out\IrrWaveDKTest\";
Path="X:\analysis\projects\E2128 Reliance\ModelTest\ftp.Marin.nl\4_Results\Out\";
!Path="F:\NOV_PC\D_Drive\ModelTest\31448_Marin_Reliance2128\ftp.marin.nl\4_Results\out\";
VarList=Reshape([&
"03\005\05\","31448_01OB_03_005_005_01","2160" &
                                               ],[nLC,3],order=[2,1]);
nCol2Use=size(VarHdr)
XLfil='IrrWaveTestCalib.xls'; open(iflog,file='logcsv.txt');
Bukhdl=xlCreateBook(); iBukLoad=xlBookLoad(Bukhdl,ctrim(XLfil)//""C); !CALL DATE(DateStr)
if(iBukLoad==1)Arkhdl0=xlBookGetSheetbyName(Bukhdl,'0'C)
DO II=1,nLC
  iFil=Path//trim(VarList(II,2))//"_TimeTrace.plt" !31448_01OB_03_ 003_001_01_UF.csv
  !trim(VarList(II,3))
 IF(FILE_EXIST(iFil))THEN
  open(iui,file=iFil,status='old'); ios=0; ncol=0; iRow=0
      ios=lineread(iui,FILOUT,ipos); if(ios==-1)exit
      k=index(FILOUT,'Variables='); if(k>0)FILOUT=FILOUT(k+1:size(FILOUT))
      k=index(FILOUT,'ZONE'); if(k>0)FILOUT=FILOUT(:k-1)
      call reallocate(ipos,size(FILOUT))
      call LineParser(FILOUT,ipos,strArr,numArr)
      nStr=size(strArr); nNum=size(numArr); call reallocate(iNewOrd,nStr); iNewOrd=[(i,i=1,nStr)]; call reallocate(iColmn,nCol2Use)
      if(nNum==0.and.nStr>nCol2Use)then
       do k=1,nCol2Use
        j=index(strArr(iNewOrd),ctrim(VarHdr(k)));
        if(j==0)then
         call errormsg(iFil//' does not contain Channel:'//ctrim(VarHdr(k))//'!',-1); cycle
        end if 
        if(iNewOrd(k)>0)then
          call Swap(iNewOrd(j),iNewOrd(k))  
        elseif(iNewOrd(k)==0)then
          iNewOrd(k)=k
        end if  
       end do
      else
       call errormsg(iFil//' contains invalid input numbers!',-1)
      end if; nNum=nStr ! ios=lineread(iui,crad,ncol);
   iRow=iRow+1; call reallocate(Table,26000,nNum)
   do while(ios/=-1) !read iMea=iRow      
     read(iui,*,iostat=ios)Table(iRow,:)
     if(ios==0)then; !if(iRow>20)then; call reallocate(Table,iRow,nNum); exit; end if
      iRow=iRow+1;
      if(iRow>26000)then
        call reallocate(Table,iRow,nNum)
      end if  
     elseif(ios==-1)then
      iRow=iRow-1; call reallocate(Table,iRow,nNum)
      do i=1,iRow
        Table(i,:)=Table(i,iNewOrd);
      end do
      exit;
     else
      call errormsg(iFil//' contains invalid input number format at line:'//num2str(iRow+1)//'->expecting'//num2str(nNum)//' columns!',-1)
     end if
   end do;  close(iui); 
   ArkNam=trim(VarList(II,3))
   crad=ArkNam//'.txt'; call FOpenFile(iuo,crad); nCh2Asc=25 !15
   write(iuo,'(I)')iRow
   !call reallocate(NumArr,2); NumArr(1)=mean(Table(2:,1)-Table(:iRow-1,1)) !mean timestep
   !NumArr(2)=Table(2,1)-Table(1,1) !timestep
   if(abs(mean(Table(2:,1)-Table(:iRow-1,1))-(Table(2,1)-Table(1,1)))>1d-1)then
     call ErrorMsg(iFil//' contains variant time steps ->expecting constant time step!',-1)
   end if
   write(iuo,'(G10.5)')Table(2,1)-Table(1,1)
   write(iuo,'(I)')nCh2Asc;
   !call reallocate(ipos,nCh2Asc); iPos=[(i,i=2,10),15,16,18,22,23,24]
   iPos=[(i+1,i=1,nCh2Asc)]![(i,i=2,10),15,16,18,22,23,24]
   do j=1,nCh2Asc
    write(iuo,'(a)')trim(strArr(iNewOrd(iPos(j))))
   end do
   do i=1,iRow
    write(iuo,'(<nCh2Asc-1>(G13.5,1x),G13.5)')Table(i,iPos)
   end do; close(iuo)
   cycle
   write(iflog,'(a)')iFil//' has been read successfully @ '//fdate()
   if(iBukLoad==1)Arkhdl=xlBookGetSheetbyName(Bukhdl,ArkNam//''C)
   if(Arkhdl%Point==0)Arkhdl=xlBookAddSheet(Bukhdl,ArkNam//''C,Arkhdl0)
   if(Arkhdl%point/=0)then
     call xlSheetSplit(Arkhdl,6,1)   
     call WriteCell(Arkhdl,0,0,['Min','Max','Mean','Std'],rowcol='r')
     do j=1,nCol2Use
       !call WriteCell(Arkhdl,0,j,minval(Table(:,j+1)));
       !call WriteCell(Arkhdl,1,j,maxval(Table(:,j+1)));
       !call WriteCell(Arkhdl,2,j,mean(Table(:,j+1)));
       !call WriteCell(Arkhdl,3,j,std(Table(:,j+1)));    
       crad=trim(XLCol2Str(j+1))
       i_ret=xlSheetWriteFormula(Arkhdl,0,j,"=MIN("//crad//"7:"//crad//num2str(6+iRow)//")"C); 
       i_ret=xlSheetWriteFormula(Arkhdl,1,j,"=MAX("//crad//"7:"//crad//num2str(6+iRow)//")"C); 
       i_ret=xlSheetWriteFormula(Arkhdl,2,j,"=AVERAGE("//crad//"7:"//crad//num2str(6+iRow)//")"C); 
       i_ret=xlSheetWriteFormula(Arkhdl,3,j,"=STDEV("//crad//"7:"//crad//num2str(6+iRow)//")"C); 
       i_ret=xlSheetWriteFormula(Arkhdl,4,j,"=COLUMN()"C); 
     end do
     do i=6,iRow+6
       i_ret=xlSheetWriteFormula(Arkhdl,i,nCol2Use,"=SQRT(SUMSQ(O"//num2str(i+1)//"+XTUR;P"//num2str(i+1)//"))"C);
     end do 
     call WriteCell(Arkhdl,5,0,strArr(iNewOrd(:nCol2Use))); call reallocate(Table,iRow,nCol2Use)
     call WriteCell(Arkhdl,6,0,Table);
     i_ret=xlBookSave(Bukhdl, XLfil//""C); Arkhdl%Point=0
   elseif(Arkhdl%Point==0)then
     call ErrorMsg('Error in loading sheet:'//ArkNam//' of '//XLfil,-1)
   end if
 END IF
END DO; close(iflog)
!call xlBookRelease(Bukhdl); 
End subroutine IrrWaveTest
    
subroutine WeiBullFitXLData(XLfil,ArkNam,TransT)
use util_xk,only:reallocate,swap,mean,std
use LStrFileUti
use libxl
use StatPack,only:WeiBulTailExtreme!,LevelCrossing,StatPara,WeibullMLE4GB,WeibullMME_Skewness4B,FITDIS
IMPLICIT  NONE
character(*),intent(in)::XLfil,ArkNam
real,intent(in)::TransT
type(BookHandle)::Bukhdl0
type(SheetHandle)::Arkhdl0
integer::nfmt,nfont,iBukLoad,iBukOld,i_ret,itype
character(:),allocatable::crad,filnam,ArkNam0,ErrMsg
integer::iui=1,iflog=2,iuo=3,i,ir,II,j,jc,k,m,ios,size,nRec,nStep,nHdr,iRow,iCol,nCol,nTr,nCol2Use
real*8,DIMENSION(:,:),allocatable::Table,xEXt
real*8::Twind,dum
Bukhdl0=NEWXMLBook(); 
iBukLoad=xlBookLoad(Bukhdl0,ctrim(XLfil)//''C);
!iBukload=xlBookLoadPartiallyUsingTempFile(Bukhdl0,ctrim(Dir)//ctrim(XLfil)//""C,0,0,125904,tempFile,1)
IF(iBukLoad==1)then; 
   Arkhdl0=xlBookGetSheetbyName(Bukhdl0,ArkNam//''C)
   if(Arkhdl0%Point==0)then
       call ErrorMsg('Error in loading sheet:'//ArkNam//' of '//XLfil,-1)
   else
!       Now the Dest.Excel is ready, just need to read the Source data from ArkHdl0
     write(*,'(a)')'Now processing: '//XLfil
        iRow=0; i=iRow; i_ret=xlSheetCellType(Arkhdl0,i,0);
        do while(i_ret/=CELLTYPE_NUMBER)
         i=i+1; i_ret=xlSheetCellType(Arkhdl0,i,0); if(i_ret==0)i_ret=xlSheetCellType(Arkhdl0,i,1)
        end do; nHdr=i

        do while(i_ret==CELLTYPE_NUMBER)
         i=i+1; i_ret=xlSheetCellType(Arkhdl0,i,0)
         if(i_ret==CELLTYPE_STRING.or.i_ret==CELLTYPE_EMPTY.or.i_ret>=CELLTYPE_BLANK)then; 
             iRow=i-1; exit;     !0-based Row Number = i - 1
         end if
        end do !CELLTYPE_EMPTY=0,CELLTYPE_NUMBER=1,CELLTYPE_STRING=2,CELLTYPE_BOOLEAN=3,CELLTYPE_BLANK=4,CELLTYPE_ERROR=5
         j=1; i_ret=xlSheetCellType(Arkhdl0,iRow,j); nRec=iRow+1-nHdr
         do while(i_ret==CELLTYPE_NUMBER)
          j=j+1; i_ret=xlSheetCellType(Arkhdl0,iRow,j)
          if(i_ret==CELLTYPE_STRING.or.i_ret==CELLTYPE_EMPTY.or.i_ret>=CELLTYPE_BLANK)then; 
              nCol2Use=j; exit;        !0-based Number = j - 1
          end if
         end do
        if(allocated(Table))deallocate(Table); Allocate(Table(nRec,nCol2Use),stat=ios)!call reallocate(Table,iRow-2,nCol2Use)
        i=1;
        do while(ios>0)
          Allocate(Table(nRec,nCol2Use),stat=ios)
          call ErrorMsg('Error in Memory allocation: Alloc_ERR='//num2str(ios)//'!Retry#'//num2str(i),2*i)
        end do
        if(allocated(xExt))deallocate(xExt); Allocate(xExt(2,nCol2Use-1))
        DO ir=nHdr,iRow
          do jc=0,nCol2Use-1
            Table(ir-nHdr+1,jc+1)=xlSheetReadNum(Arkhdl0,ir,jc);
           end do
        END DO; Twind=Table(nRec,1)-Table(1,1)-TransT; nTr=int(TransT/(Table(2,1)-Table(1,1)))+1
        do jc=1,nCol2Use-1; dum=maxval(Table(nTr:,jc+1))
            if(dum<0)then! For all negative Fz
              call WeiBulTailExtreme(-Table(nTr:,jc+1),Twind,xExt(:,jc)); xExt(1,jc)=-xExt(1,jc)
            else
              call WeiBulTailExtreme(Table(nTr:,jc+1),Twind,xExt(:,jc))
            end if  
         call WriteCell(Arkhdl0,iRow+5,jc,xExt(:,jc),rowcol='Row')
        end do
      i_ret=xlBookSave(Bukhdl0,ctrim(XLfil)//''C)
      call xlBookRelease(Bukhdl0); 
   end if 
ELSE
   ErrMsg=GetErrorMsg(Bukhdl0);
   call ErrorMsg(ErrMsg//' in loading Existing summary file:'//ctrim(XLfil),-1)
END IF
end subroutine WeiBullFitXLData    