Module StatPack
!  Symbolic names for kind types of 4-, 2-, and 1-byte integers:
	INTEGER, PARAMETER::I4B = SELECTED_INT_KIND(9)
!  Symbolic names for kind types of single- and double-precision reals:
    INTEGER, PARAMETER::DP = KIND(1D0)
	REAL(DP),PARAMETER::ERROR=1d-10
    
    INTERFACE swap;  MODULE PROCEDURE swap_i,swap_d; END INTERFACE
    INTERFACE ASSERT_EQ;  MODULE PROCEDURE ASSERT_EQ2,ASSERT_EQ3; END INTERFACE
    INTERFACE LSqErr;  MODULE PROCEDURE LSqErr0,LSqErr1; END INTERFACE
    INTERFACE reallocate;  MODULE PROCEDURE ReAllocate_IV,ReAllocate_dV,ReAllocate_IM,ReAllocate_DM; END INTERFACE

    contains
    FUNCTION ASSERT_EQ2(N1,N2,STRING)
    CHARACTER(LEN=*),INTENT(IN)::STRING
    INTEGER,INTENT(IN)::N1,N2
    INTEGER::ASSERT_EQ2
    IF (N1 == N2) THEN
    ASSERT_EQ2=N1
    ELSE
    WRITE (*,*) 'NRERROR: AN ASSERT_EQ FAILED WITH THIS TAG:',STRING; Pause
    STOP 'PROGRAM TERMINATED BY ASSERT_EQ2'
    END IF
    END FUNCTION ASSERT_EQ2
    
    FUNCTION ASSERT_EQ3(N1,N2,N3,STRING)
    CHARACTER(LEN=*),INTENT(IN)::STRING
    INTEGER,INTENT(IN)::N1,N2,N3
    INTEGER::ASSERT_EQ3
    IF (N1 == N2 .AND. N2 == N3) THEN
    ASSERT_EQ3=N1
    ELSE
    WRITE (*,*) 'NRERROR: AN ASSERT_EQ FAILED WITH THIS TAG:',STRING; Pause
    STOP 'PROGRAM TERMINATED BY ASSERT_EQ3'
    END IF
    END FUNCTION ASSERT_EQ3
    
    subroutine ReAllocate_IV(P,N,n1)
    INTEGER(I4B),DIMENSION(:),allocatable::P,R
    INTEGER(I4B),INTENT(IN)::N
    INTEGER(I4B),INTENT(IN),optional::n1
    INTEGER(I4B)::IERR,i,i1,iL,iU
    i1=1; if(present(n1))i1=n1; 
    ALLOCATE(R(i1:i1+N-1),STAT=IERR); R=0
    IF (IERR /= 0) WRITE(*,'(a)')'REALLOCATE_IV: PROBLEM IN ATTEMPT TO ALLOCATE MEMORY'
    IF(allocated(p))then
    iL=LBound(P,1); iU=UBound(P,1); 
     do i=iL,min(iU,n-i1+1)
      R(i1+i-iL)=P(i)
     end do
    end if
    call move_alloc(R,P)
    END subroutine ReAllocate_IV
    
    subroutine ReAllocate_dV(P,N,n1)
    REAL(dP),DIMENSION(:),allocatable::P,R
    INTEGER(I4B),INTENT(IN)::N
    INTEGER(I4B),INTENT(IN),optional::n1
    INTEGER(I4B)::IERR,i,i1,iL,iU
    i1=1; if(present(n1))i1=n1; 
    ALLOCATE(R(i1:i1+N-1),STAT=IERR); R=0d0
    if (ierr /= 0) write(*,'(a)')'REALLOCATE_dV: problem in attempt to allocate memory'
    if(allocated(p))then
    iL=LBound(P,1); iU=UBound(P,1); 
     do i=iL,min(iU,n-i1+1)
      R(i1+i-iL)=P(i)
     end do
    end if
    call move_alloc(R,P)
    END subroutine ReAllocate_dV
    
    subroutine ReAllocate_dM(P,N,M,n1,m1)
    REAL(dP),DIMENSION(:,:),allocatable::P,R
    INTEGER(I4B),INTENT(IN)::N,M
    INTEGER(I4B),INTENT(IN),optional::n1,m1
    INTEGER(I4B)::IERR,i,j,i1,j1,iL,iU,jL,jU
    i1=1; j1=1; if(present(n1))i1=n1; if(present(m1))j1=m1
    ALLOCATE(R(i1:i1+N-1,j1:j1+M-1),STAT=IERR); R=0d0
    IF (IERR /= 0) WRITE(*,'(a)')'REALLOCATE_dM: PROBLEM IN ATTEMPT TO ALLOCATE MEMORY'
    IF(allocated(p))then
    iL=LBound(P,1); iU=UBound(P,1); 
    jL=LBound(P,2); jU=UBound(P,2); 
     do i=iL,min(iU,n-i1+1)
       do j=jL,min(jU,m-j1+1)
        R(i1+i-iL,j1+j-jL)=P(i,j)
       end do
     end do
    end if
    call move_alloc(R,P)
    END subroutine ReAllocate_dM
    
    subroutine ReAllocate_IM(P,N,M,n1,m1)
    INTEGER(I4B),DIMENSION(:,:),allocatable::P,R
    INTEGER(I4B),INTENT(IN)::N,M
    INTEGER(I4B),INTENT(IN),optional::n1,m1
    INTEGER(I4B)::IERR,i,j,i1,j1,iL,iU,jL,jU
    i1=1; j1=1; if(present(n1))i1=n1; if(present(m1))j1=m1
    ALLOCATE(R(i1:i1+N-1,j1:j1+M-1),STAT=IERR); R=0
    IF (IERR /= 0) WRITE(*,'(a)')'REALLOCATE_IM: PROBLEM IN ATTEMPT TO ALLOCATE MEMORY'
    IF(allocated(p))then
    iL=LBound(P,1); iU=UBound(P,1); 
    jL=LBound(P,2); jU=UBound(P,2); 
     do i=iL,min(iU,n-i1+1)
       do j=jL,min(jU,m-j1+1)
        R(i1+i-iL,j1+j-jL)=P(i,j)
       end do
     end do
    end if
    call move_alloc(R,P)
    END subroutine ReAllocate_IM
    
    SUBROUTINE SWAP_I(A,B)
    INTEGER(I4B),INTENT(INOUT)::A,B
    INTEGER(I4B)::DUM
    DUM=A
    A=B
    B=DUM
    END SUBROUTINE SWAP_I
    
    SUBROUTINE SWAP_D(A,B)
    REAL(dP),INTENT(INOUT)::A,B
    REAL(dP)::DUM
    DUM=A
    A=B
    B=DUM
    END SUBROUTINE SWAP_D
    
integer(4) function Get_Unit(iu_start,iu_max)
!   Get_Unit returns a unit number that is not in use
   integer(4),intent(in),optional::iu_start,iu_max
   integer(4)::i,ios,i_s,i_e
   logical::lopen,pres_iu_start,pres_iu_max
   pres_iu_start=present(iu_start); pres_iu_max=present(iu_max)
   i_s=1; i_e=9999
   if(pres_iu_start)then
    if(abs(iu_start)>i_e)then
     i_s=1
    elseif(iu_start<0)then
      i_s=-iu_start+1
    else
      i_s=iu_start
    end if
   end if
   if(pres_iu_max)then
     if(iu_max>0)i_e=iu_max
   end if    
   do i=i_s,i_e
    if(i/=5 .and. i/=6 .and. i/=9)then
	inquire(unit=i,opened=lopen,iostat=ios)
    if(ios==0.AND. .not.lopen)then;	Get_Unit=i; exit;	end if
    end if
   end do
end function Get_Unit
!::::::::::::::::::::: SUBROUTINE OR FUNCTION :::::::::::::::::::::::::::::::::::::::
!_____________________ SUBROUTINE SSort1D_D(X,Y,IPOS,N,NK,KICKSAME)______________________
    SUBROUTINE Sort(X,IPOS,NK,KICKSAME)
    IMPLICIT NONE
    INTEGER,INTENT(INOUT)::IPOS(:),NK
    LOGICAL,INTENT(IN),OPTIONAL::KICKSAME
    REAL(DP),INTENT(INOUT)::X(:)
    INTEGER::M,I,ISWAP(1),ISWAP1
    	!X(M)--:ORIGINAL DATA;	IPOS--:ORIGINAL DATA INDEX ORDER
    	!NK--:INPUT(0: unchange; >0: INCREASE SORT; <0: DECREASE SORT) & OUTPUT NK FOR DATA SIZE OF
    	!KICKED THE REPEATED PNTS
    M=size(X)
    DO I=1,M-1
         IF(NK==0)THEN
    	  Exit
         ELSEIF(NK>0)THEN
    	   ISWAP=MINLOC(X(I:M))
    	 ELSEIF(NK<0)THEN
    	   ISWAP=MAXLOC(X(I:M))
    	 END IF
    	   ISWAP1=ISWAP(1)+I-1
             IF(ISWAP1/=I) THEN
             call swap(X(I),X(ISWAP1))
!           XTMP=X(I);		X(I)=X(ISWAP1);		X(ISWAP1)=XTMP
             call swap(IPOS(I),IPOS(ISWAP1))
!           ITEMP=IPOS(I);	IPOS(I)=IPOS(ISWAP1);   IPOS(ISWAP1)=ITEMP
             ENDIF
    END DO
    	
    	IF(PRESENT(KICKSAME))THEN
    		IF(KICKSAME)THEN
    			NK=0
    			DO I=1,M-1
    			IF(ABS(X(I+1)-X(I))>ERROR)THEN
    			 NK=NK+1;	X(NK)=X(I);	IPOS(NK)=IPOS(I)
    			END IF
    			END DO
    			
    			IF(NK<M)THEN
    			 NK=NK+1;	X(NK)=X(I);	IPOS(NK)=IPOS(I)
    			END IF
    		END IF
    	END IF
    END SUBROUTINE Sort
    
!::::::::::::::::::::: SUBROUTINE OR FUNCTION :::::::::::::::::::::::::::::::::::::::
!_________function lineread0(filno,cargout,ncol)result(ios)__________________________
    function lineread(filno,cargout,ncol,delim,cont)result(ios)
! this subroutine reads 
! 1. following row in a file except a blank line or the line begins with a !#*
! 2. the part of the string until first !#*-sign is found or to end of string 
!       
! input Arguments:
!  filno(integer)      input file number                          
!  sep(character,optional)input separator
!  ncol(INOUT:integer,optional)number of columns,if ncol>=0,strlen if ncol<0
!  delim(character*)delimiters
!  cont(character*)continuation sign
! output Arguments:
!  cargout(character)  output string,converted so that all unecessay spaces/tabs/control characters removed.
!  ios(integer)reading iostat
! eg. nl=0; ios=lineread(iuin,crad,nl,' ','')output the original linesl
    implicit none
    integer,intent(in)::filno
    character(len=:),allocatable,intent(out)::cargout
    integer,intent(inout),optional::ncol
    character(*),intent(in),optional::delim
    character(*),intent(in),optional::cont
    logical::pres_ncol,pres_delim,pres_cont,origin=.false.,escape=.false.
    integer::nlen,i,ip,ich,isp,nsp,size,ios,icon,ieq,iqu,msp,iesc
    character::ch
    character,parameter::tab=char(9)
    character(:),allocatable::sep
    pres_ncol=present(ncol); 
    pres_delim=present(delim); sep='=,;()[]{}*~'; if(pres_delim)sep=delim
    pres_cont=present(cont); icon=38; if(pres_cont)icon=ichar(cont)
!=_61	,_44	;_59	(_40	)_41	[_91	]_93    ^_94	{_123	}_125	*_42	~_126	"_34	#_35    &_38    '_39
    if(len(sep)+len(cont)==0)origin=.true.
    size=0; nlen=0; isp=1; nsp=0; ich=-1; ios=0; ieq=0; iqu=0; msp=0; iesc=0; cargout=''; 
    Do While(ios/=-1)! The eof()isn't standard Fortran.
! =-2 if an end-of-record condition occurs with nonadvancing reads.
! =-1 if an end-of-file condition occurs.
     READ(filno,"(A)",ADVANCE='NO',SIZE=size,iostat=ios,ERR=9,END=9)ch 
      ich=iachar(ch); if(.not.escape)ip=scan(ch,sep); 
      if(size==0.or.ios<0)then
       READ(filno,"(A)",ADVANCE='no',SIZE=size,iostat=ios,EOR=9); if(nlen>0.or.origin)exit
      end if
      if(ich==94)then; 
         iesc=iesc+1; 
         if(mod(iesc,2)==1)then
           escape=.true.; cycle;
         else
           escape=.false.;   
         end if
      end if; iesc=0
      if(ich<=32.or.ip==1)then ! tab(9),space(32),"(34),=(61)... characters
        if(origin)then
           nlen=nlen+1; cargout=cargout//ch;
        end if
        if(isp==0)isp=1; 
        if(ieq==0.and.ich==61)ieq=61
      elseif(.not.origin.and..not.escape.and.(ich==33.or.ich==35.or.ich==icon))then !if char is comment !# or continue sign &(38)
       READ(filno,"(A)",ADVANCE='yes',iostat=ios)ch; if(nlen>0.and.ich/=icon)exit;
      else
       if(ich==34.and.iqu>0)iqu=iqu+1;
        if(nlen>0)then
         if(ieq==61)then
          nsp=nsp+1; if(.not.origin)then; nlen=nlen+1; cargout=cargout//'='; end if
         elseif(isp==1)then
          nsp=nsp+1; if(.not.origin)then; nlen=nlen+1; cargout=cargout//' '; end if  ! avoid the case for space inside dbl quotes
         end if
        end if
        if(iqu>0)then
         if(mod(iqu,2)==1)then
           if(ich==34)then
            msp=0; origin=.true.
           elseif(ieq==61.or.isp==1)then
            msp=msp+1;
           end if
         elseif(mod(iqu,2)==0)then
           nsp=nsp-msp; msp=0; origin=.false.
         end if
        elseif(iqu==0.and.ich==34)then
          origin=.true.  
        end if
        if(ich==34.and.iqu==0)iqu=iqu+1;
        nlen=nlen+1; cargout=cargout//ch; 
        if(nsp==0)nsp=1;      
        isp=0; ieq=0
      end if; escape=.false.;
    end do
    9 if(size*ios>0)type*,'Met error in reading file in [lineread]'
! ios<0: Indicating an end-of-file or end-of-record condition occurred. 
     if(ieq==61)then
        nlen=nlen+1; cargout=cargout//'=';
     end if   
    
     if(pres_ncol.and.nlen>0)then
       if(ncol<0)then; ncol=nlen; else; ncol=nsp; end if
     end if
    if(nlen>0)ios=0;
!  type*,cargout,nlen,nsp
    end function lineread
    
    function LSqErr0(iWei,X,Neg1mLogP,SHA)result(Err)
! Common Data
! nx,X(nx),P(nx)
     integer,intent(in)::iWei(:)
     real*8,intent(in)::X(:),Neg1mLogP(:),SHA
     real*8::Err
     integer::Nx,NWei
     real*8::Sx,Sy,Sxx,Sxy,Syy
     real*8::Y(size(X))
! Nx=ASSERT_EQ(size(iWei),size(X),size(P),'Input to LSqErr(X,P... has size error!')! Comment to be faster, Uncomment to check
     Y=(Neg1mLogP)**(1d0/SHA)!(-LOG(1d0-P))**
     SX=sum(X*iWei)
     SY=sum(Y*iWei)
     SXX=sum(X*X*iWei)
     SXY=sum(X*Y*iWei)
     SYY=sum(Y*Y*iWei)
     NWei=sum(iWei)
     ERR=sum(iWei*(X-(NWei*Sxy-Sx*Sy)/(NWei*Syy-Sy*Sy)*Y-(Sx-(NWei*Sxy-Sx*Sy)/(NWei*Syy-Sy*Sy)*Sy)/NWei)**2)
     end function LSqErr0
    
    function LSqErr1(Nx,iWei,X,Neg1mLogP,SHA)result(Err)
! Common Data
! nx,X(nx),P(nx)
     integer,intent(in)::Nx
     integer,intent(in)::iWei(Nx)
     real*8,intent(in)::X(Nx),Neg1mLogP(Nx),SHA
     real*8::Err
     integer::i,NWei
     real*8::Sx,Sy,Sxx,Sxy,Syy
     real*8::Y(Nx)
     Y=(Neg1mLogP)**(1d0/SHA)!(-LOG(1d0-P))**
     SX=sum(X*iWei)
     SY=sum(Y*iWei)
     SXX=sum(X*X*iWei)
     SXY=sum(X*Y*iWei)
     SYY=sum(Y*Y*iWei)
     NWei=sum(iWei)
     ERR=sum(iWei*(X-(NWei*Sxy-Sx*Sy)/(NWei*Syy-Sy*Sy)*Y-(Sx-(NWei*Sxy-Sx*Sy)/(NWei*Syy-Sy*Sy)*Sy)/NWei)**2)
    end function LSqErr1
    
    FUNCTION Mean(X,ni,n) ! return the mean value or Expected Maximum
    REAL(dP),DIMENSION(:),INTENT(IN)::x
    integer,INTENT(IN),optional::ni(:)
    integer,INTENT(IN),optional::n
    REAL(dP)::Mean
    integer::nasrt
    logical::ni_pres,n_pres
    ni_pres=present(ni); n_pres=present(n)
    if(ni_pres)then
        nasrt=assert_eq(size(x),size(ni),'Mean X and ni has different dimension!')
        Mean=sum(x*ni)/dble(sum(ni)) ! xi: mid-point in class number i; ni: number of samples in class number i
     if(n_pres)Mean=sum(ni*(x-Mean_dp)**n)/dble(sum(ni))
    else
        Mean=sum(x)/dble(size(x))
    end if
    END FUNCTION Mean
    
    FUNCTION STD(X,flag) ! return the standard deviation
    REAL(dP),DIMENSION(:),INTENT(IN)::x
    integer,INTENT(IN),optional::flag
    REAL(dP)::STD
     if(present(flag).and.flag==1)then
      STD=sqrt(sum((x-mean(x))**2)/dble(size(x)))
     else
      STD=sqrt(sum((x-mean(x))**2)/dble(size(x)-1))
     end if
    END FUNCTION STD
    
subroutine WeiBulTailExtreme(X,Twind,xExt,pValue,WeibPara,VarName)
use kernel32,only:MoveFileEx,CreateDirectory,NULL
implicit none
real(8),intent(in)::X(:),Twind
real(8),intent(out)::xExt(2)
real(8),intent(in),optional::pValue
real(8),intent(out),optional::WeibPara(3)
character(*),intent(in),optional::VarName
integer,parameter::nMom=4,nCLS=10; 
real*8,DIMENSION(:),allocatable::XParent,XMaxima,XMinima,Pr,pdf,CDF
real*8::LOC_OLD,SCA_OLD,SHA_OLD,PLKey,P_ext,StatPar(6),CLsz,gama,beta,mu,xMomc(0:nMom),GaBeMu(3),MLEpar(1),MMEPar(2)
character(:),allocatable::fName,xTitle,yTitle,Dir,VarNm
integer::pUnique,ErrType,iWei,nMxNo,i
!pValue: Statistic percentile=>any one of the hundred groups so divided while fractile is the value of a distribution for which some fraction of the sample lies below.
real*8::pVal=.37d0,pLow=.80d0,MxTmp!=PFRAC; else; pValue=.37d0
real,parameter::ThreeHr=10800!=PFRAC; else; pValue=.37d0
!if(pres_PLKey)then; pLow=PLKey; else; pLow=.80d0; endif
! The following is to build up the statistic probility distribution model for WeiBull fitting
!     For Long-term 100yr ULS Extreme: F(x)^N=1-1/100=>F(x)=(1-.01)^(1/N); Annual probability of non-exceedance for N (independent events per year), e.g.N=20/35(20 storms in 35 years)
!     For short-term 3-hr SS Extrem:F(x)^N=PFRAC=>F(x)=PFRAC^(1/N);Non-exceedance PFRAC for N peaks and no# of 3hr,=PFRAC**(TimeWinUpp-TimeWinLow)/(NoPeaks*PERIOD)          
!     The exceedance probability 1/N of a single 3-hour maximum to exceed the value for the limit states are: 1/(100*2922) and 1/(10000*2922) for 100yr-ULS and 10000yr-ALS respectively.
!Dir="Weibull_Fit_data\"; i=CreateDirectory(Dir//""C,NULL)
if(present(pValue))then; pVal=pValue; else; pVal=.37d0; end if
if(present(VarName))then; VarNm=Dir//trim(VarName); else; VarNm=Dir; end if
call StatPara(X,nCLS,StatPar,VarNm//'ParentSeries.txt')!StatPar=[rMin,xMean,rMax,rStd,rSkew,rKurt]
call LevelCrossing(X,XMaxima); nMxNo=size(XMaxima); Pr=[(dble(i)/(nMxNo+1d0),i=1,nMxNo)]; P_ext=pVal**(Twind/(nMxNo*ThreeHr))
call StatPara(xMaxima,nCLS,StatPar,VarNm//'MaxSeries.txt'); LOC_OLD=StatPar(1)-1d0/nMxNo; SCA_OLD=0d0; SHA_OLD=0d0
!     call reallocate(xMaxima,nPnt); xMaxima=X; Pr=XX(ir+2,:); P_ext=.99999d0
call FITDIS(xMaxima,Pr,P_ext,pLow,pUnique,ErrType,iWei,xExt,LOC_OLD,SCA_OLD,SHA_OLD,yTitle=VarNm); 
WeibPara=[LOC_OLD,SCA_OLD,SHA_OLD]
!close(114); i=MoveFileEx('fort.114',VarNm//'FitResp.txt'C,0)

end subroutine WeiBulTailExtreme


    SUBROUTINE StatData(infil,xlBuk)
!use libxl
    implicit none
    character*(*),intent(in)::infil
!type(BookHandle),intent(in)::xlBuk
    integer,intent(in)::xlBuk
      real*8,DIMENSION(:),allocatable::X,xd,XParent,XMaxima,XMinima,Pr,pdf,CDF
      real*8,DIMENSION(:,:),allocatable::XX,PP,Xex
      integer::i,ir,j,k,L,m,p,iuin,nResp,nPnt,ios,nc,nMxp,iloc(1),NK,nMxNo
      INTEGER,DIMENSION(:),allocatable::indcr0,induc,inddc,indPo,indNe,iPos,iPosp,iPosn,locmx,nMxi
      INTEGER::mdp,mdn,mlv
      REAL*8::lv,lvband,lvsaddle,gama,beta,mu,  dum1,dum2,dum3
      integer::N1,N2,IpStart,isUnique,ErrorKey,iWeight,IEXT,nCLS,ipo=0,ine=0,jmp=0,mit=0
      integer,parameter::nMom=4
      real*8::LOCD,SCAD,SHAD,PLKey,P_ext,rMax,rMin,rMean,xMean,rStd,rSkew,rKurt,CLsz,xMomc(0:nMom),GaBeMu(3),MLEpar(1),MMEPar(2),StatPar(6)
      character(:),allocatable::crad,ArkNam(:),fName(:),xTitle(:),yTitle
      logical::check_return
!ArkNam=["T1","T2","T3","T4","T5","T6","T7","T8","T9","T10","Fxy","Sxy","R","V","H","Fzs"]  
!fName=['time04.totDistFit.plt',&
! xTitle=['Tension line  1 [kN]       ',&  
yTitle="-Ln[1-Prob(X<math>#</math>x)]"
!   P_ext=(1d0-.01d0)**dble(35d0/20d0) ! 20 storms in 35 years
!Pr=EULER!0.367879441d0=1d0/2.71828182845904d0
      PLKey=.01d0; isUnique=1; ErrorKey=2; iWeight=1; iuin=193
      crad=infil; open(iuin,file=crad,status='old'); ios=0; nPnt=0; nResp=0
     do while(ios/=-1)
       nc=0; ios=lineread(iuin,crad,nc); 
       if(verify(crad,' DdEe+-.0123456789')>0)cycle
       if(ios==0)then
        nPnt=nPnt+1
         if(nc>1)call reallocate(XX,nc,nPnt)
         read(crad,*,iostat=ios)XX(:,nPnt)
         if(nResp>0.and.nResp/=nc-1)write(*,'(a,i)')infil//' missing response vars for PointNo#',nPnt
         nResp=nc-1 
       end if
     end do
       N1=nResp; m=nPnt; allocate(PP(N1,m),Xex(N1,2)); !forall(i=1:nResp)PP(i,:)=[(dble(i)/(m+1d0),i=1,m)]
      allocate(X(m)); nCLS=10; 
      DO ir=1,nResp !global NoiResp in RootFinding
       X=XX(ir+1,:);! Pr=PP(ir+1,:)
       call LevelCrossing(X,XMaxima)
       call StatPara(xMaxima,nCLS,StatPar)
!Initialize gamma,beta,mu: A better estimate of mu=xMin−1/n (see Sirvanci and Yang, 1984, p. 74)
       mu=Minval(xMaxima)-1d0/size(xMaxima)
!MLEpar=[0.653486457990338d0]!,3.219113441547919d-5]!,rMin*.99d0]!rMin-1d0/nMxNo]
!MMEpar=1d0![1.1222d0,1.5d2];
       GaBeMu=[1d0,0d0,mu]
!   do
!          dum1=FunMLEGetG(gama)
!          dum2=FunMLEGetG(gama*1.01d0)
!          if(dum2>dum1)then
!           if(ine>1)then; ine=0; jmp=jmp+1; else; jmp=0; endif
!           ipo=ipo+1
!          else
!           if(ipo>1)then; ipo=0; jmp=jmp+1; else; jmp=0;  endif
!           ine=ine+1
!          end if
!          write(*,'(g0,3(1x,g0),3i)')gama,dum1,dum2,(dum2-dum1)/(gama*.01),ipo,ine,jmp! 0.718480731860414   79705.2541164183   79705.2541164183 0.10127E-08
!          gama=gama-dum1/((dum2-dum1)/(gama*.01)) 
!          !Here Newton-R's method won't work due to non-zero Error. Gradient scaled and can be improved to a function of numbers of downsliding and uprising in case some hard situation.
!     if(isnan(dum1))then
!        write(*,*)gama,dum1,dum2 
!     end if    
!          if(abs(dum1)<1d-10)exit
!   end do
       call WeibullMLE4GB(xMaxima,GaBeMu)
       gama=1d0; call WeibullMME_Skewness4B(rSkew,gama); beta=rStd/sqrt(Gamma(1d0+2d0/gama)-Gamma(1d0+1d0/gama)**2)
       write(*,*)'MLEpar:',GaBeMu
       write(*,*)'MMEpar:',gama,beta
    
     if(xlBuk>0)call FITDIS(X,Pr,P_ext,PLKey,isUnique,ErrorKey,iWeight,Xex(i,:),LOCD,SCAD,SHAD,fName(i),xTitle(i),yTitle,xlBuk,ArkNam(i))
!  close(114); call system('move/y fort.114 FitResp'//num2str(i)//'.txt')
      END DO !ir,nResp
    END subroutine StatData
    
    subroutine XYSumm(iWei,X,Y,SX,SY,SXX,SXY,SYY)
! Common Data
! nx,X(nx),P(nx)
     integer,intent(in)::iWei(:)
     real*8,intent(in)::X(:),Y(:)
     real*8,intent(out)::SX,SY,SXX,SXY,SYY
     real*8::XW(size(X)),YW(size(Y))
     XW=X*iWei; SX=sum(XW)
     YW=Y*iWei; SY=sum(YW)
     SXX=sum(X*XW)
     SXY=sum(X*YW)
     SYY=sum(Y*YW)
    end subroutine XYSumm
    
  subroutine LevelCrossing(X,XMaxima)
  implicit none
      real*8,DIMENSION(:),intent(in)::X
      real*8,DIMENSION(:),allocatable::xd,XParent,XMaxima,XMinima,Pr
      real*8,DIMENSION(:,:),allocatable::XX,PP,Xex
      integer::i,ir,j,k,L,m,p,iuin,nResp,nPnt,ios,nc,nMxp,iloc(1),NK,nMxNo
      INTEGER,DIMENSION(:),allocatable::indcr0,induc,inddc,indPo,indNe,iPos,iPosp,iPosn,locmx
      INTEGER::mdp,mdn,mlv
      integer,parameter::nMom=4
      real*8::rMax,rMin,rMean,xMean,rStd,rSkew,rKurt,CLsz,gama,xMomc(0:nMom),GaBeMu(3),MLEpar(1),MMEPar(2),LocationPar
      real*8::lv,lvband,lvsaddle
      m=size(x,1)
      if(allocated(xd))deallocate(xd); p=m-1; allocate(xd(p))
      rMax=maxval(x); rMin=minval(x); rStd=std(x); rMean=mean(x); ! one option is to scale the data by x=(x-mu_p)/std_p=(x-rMean)/rStd
      xd=(x(2:m)-x(1:m-1))
      lv=0d0; lvsaddle=min(-rMax,-rMin,rMin)
! xd=diff(xm); p=m-1; ! indP=find((xd(1:p-1)>=0).*(xd(2:p)<0))+1; indN=find((xd(1:p-1)<=0).*(xd(2:p)>0))+1;
   mdp=count((xd(1:p-1)>=0d0)*(xd(2:p)<0d0)*(x(:p-1)>lv+lvsaddle)); call reallocate(indPo,mdp) ! positive peaks=>crest
   !if(mdp>0)indPo=pack([(i,i=1,p-1)],(xd(1:p-1)>=0d0)*(xd(2:p)<0d0)*(x(:p-1)>lv+lvsaddle))+1; 
   if(mdp>0)then; k=0
     do i=1,p-1
       if((xd(i)>=0d0)*(xd(i+1)<0d0)*(x(i)>lv+lvsaddle))then
        k=k+1; indPo(k)=i+1; 
       end if 
     end do  
   end if
   mdn=count((xd(1:p-1)<=0d0)*(xd(2:p)>0d0)*(x(:p-1)<lv-lvsaddle)); call reallocate(indNe,mdn) ! negative peaks=>trough
   !if(mdn>0)indNe=pack([(i,i=1,p-1)],(xd(1:p-1)<=0d0)*(xd(2:p)>0d0)*(x(:p-1)<lv-lvsaddle))+1; 
   if(mdn>0)then; k=0
     do i=1,p-1
       if((xd(i)<=0d0)*(xd(i+1)>0d0)*(x(i)<lv-lvsaddle))then
        k=k+1; indNe(k)=i+1; 
       end if
     end do  
   end if
   call reallocate(xMaxima,mdp); xMaxima=x(indPo) 
   call reallocate(xMinima,mdn); xMinima=x(indNe) 
   call reallocate(iPosp,mdp); call reallocate(iPosn,mdn); 
   iPosp=[(i,i=1,mdp)]; NK=-1; nMxp=0
   call sort(xMaxima,iPosp,NK) ! sort the response from largest to smallest and make unique
   indPo=indPo(iPosp)
   iPosn=[(i,i=1,mdn)]; NK=-1;
   call sort(xMinima,iPosn,NK) ! sort the response from largest to smallest and make unique
   indNe=indNe(iPosn)
   do j=1,mdn-1
!The procedure will also establish an upper and lower treshold at the mean value of the time 
! series for the minima and maxima distributions, respectively.
! The principle is to introduce a control level above the largest maximum (a). This 
! level is lowered until the first local minimum is found (b). The maximum above 
! this level which is closest to the minimum is then excluded. The control level 
! is further lowered until the next local minimum is found (c) and the closest 
! maximum on the upper side is excluded. The control level is finally lowered 
! until the mean level of the time series and all maximum below is excluded
           iloc(1)=0
           if(j==1)then
            where(xMinima(j)<xMaxima)iloc=minloc(xMaxima-xMinima(j))
           elseif(rMean<xMinima(j))then
            where(xMinima(j)<xMaxima.and.xMaxima<xMinima(j-1))iloc=minloc(xMaxima-xMinima(j))
           end if; p=iloc(1)
              if(p==1.or.p==mdp)then ! for debug...
                !write(*,'(a,i)')"The first minima occurs at the point:",p
              elseif(1<p.and.p<mdp)then
               if((indPo(p-1)<indNE(j).and.indNE(j)<indPo(p)).or.(indPo(p)<indNE(j).and.indNE(j)<indPo(p-1)))then ! 
!local minimum is between the two maxima
                iPosp(p)=-iPosp(p) !nMxp=nMxp+1; call reallocate(locMx,nMxp); locMx(nMxp)=-p
!xMaxima(1:ndp-nMxp)=[xMaxima(1:p-1),xMaxima(1:p+1)]; 
!ndp=ndp-1; call reallocate(xMaxima,ndp)
!call swap(xMaxima(iloc(1)),xMaxima(mdp-nMxp+1))
               end if
              end if 
!if(xMinima(j+1)<x(indNe(k)).and.x(indNe(k))<xMinima(j))then
!if(xParent(j+1)<x(indNe(k)).and.x(indNe(k))<xParent(j))then
       end do;
    nMxp=count(iPosp>0.and.rMean<xMaxima); call reallocate(xd,nMxp); xd=pack(xMaxima,iPosp>0.and.rMean<xMaxima); 
    call reallocate(xMaxima,nMxp); !Flip the order to increasing
    do i=1,nMxp
     xMaxima(i)=xd(nMxp-i+1);
    end do 
    end subroutine LevelCrossing
    
    subroutine StatPara(xMaxima,nCLS,StatPar,ofname)
    real*8,DIMENSION(:),intent(in)::xMaxima
    integer,intent(in)::nCLS
    real*8,intent(out)::StatPar(6)
    character*(*),intent(in),optional::ofname
    integer,parameter::nMom=4
    real*8::pdf(nCLS),CDF(nCLS),rMax,rMin,rMean,xMean,rStd,rSkew,rKurt,CLsz,gama,xMomc(0:nMom),GaBeMu(3),MLEpar(1),MMEPar(2)
    integer::i,k,nMxi(nCLS),nMxNo
    rMax=maxval(xMaxima); rMin=minval(xMaxima); CLsz=(rMax-rMin)/nCLS; xMomc=0d0
!rMean=mean(x); ! Estimate from sample
       do K=1,nCLS-1
        nMxi(K)=count(rMin+CLsz*(K-1)<=xMaxima.and.xMaxima<rMin+CLsz*K)
        CDF(K)=sum(nMxi(1:K))
       end do !nCLS
        nMxi(K)=count(rMin+CLsz*(K-1)<=xMaxima.and.xMaxima<=rMin+CLsz*K)
        nMxNo=sum(nMxi); CDF(K)=nMxNo;  
!CDF=CDF/nMxNo! normalize the CDF to the total no of nMxi
        xMean=sum(nMxi*(rMin+CLsz*[(k-.5d0,k=1,nCLS)]))/nMxNo! Estimate from pdf in class info 
       do i=0,nMom 
        do k=1,nCLS
          xMomc(i)=xMomc(i)+(rMin+CLsz*(k-.5d0)-xMean)**i*nMxi(k)/nMxNo
        end do
       end do
! The variance of x is thus the 2nd central moment of the probability distribution when x_o is the mean value or first moment. The 1st central moment is zero when defined with reference to the mean, so that centered moments may in effect be used to "correct" for a non-zero mean.
       if(abs(xMomc(1))>1d-3)write(*,'(a)')"The 1st central moment>1d-3! which should be zero when defined with reference to the mean"
       rStd=sqrt(xMomc(2)); rSkew=xMomc(3)/sqrt(xMomc(2)**3); rKurt=xMomc(4)/xMomc(2)**2
       StatPar=[rMin,xMean,rMax,rStd,rSkew,rKurt]
   !iuo=Get_Unit(); if(present(ofname))open(iuo,file=ofname,status='unknown')
   !!write(19,'(f13.6)')(xMaxima(nMxp-i+1),i=1,nMxp)
   !!write(19,'(2(1x,es13.6))')(nMxi(i)/nMxNo,CDF(i),i=1,nCLS)
   ! !Number UpperLimit((kN))         PDF         CDF
   !write(iuo,'(a)')'#Class UpperLimit     #Pnts-pdf    #Pnts-cdf'
   !write(iuo,'((i4,2x,ES13.6,2(1x,i)))')(i,rMin+CLsz*i,nMxi(i),int(CDF(i)),i=1,nCLS)
   !write(iuo,'(a)')'#th Central Momentum'
   !write(iuo,'((i3,(1x,ES13.6)))')(i,xMomc(i),i=0,nMom)
   !write(iuo,'(6(2x,a8,4x))')'Min     ','Mean    ','Max     ','Std     ','Skewness','Kurtosis'
   !write(iuo,'(6(1x,ES13.6))')StatPar
   !close(iuo)
    end subroutine StatPara

   subroutine WeibullMLE4GB(S,X)
   implicit none
   real*8::S(:),X(:)!S: sample data, X: Weibull par.β:scale(Characteristic Life),γ:shape(Weibull Slope),μ:shift/location(Waiting Time)
   real*8::dum1,dum2,g1,g2,df,mu
   integer::nPnt
   g1=X(1); mu=X(3); nPnt=size(S);
!mu=LocationPar
   if(g1<0d0)g1=.1d0
   dum1=1d0;
   do while(abs(dum1)>1d-10)
        dum1=nPnt/g1+sum(log(S-mu))-nPnt*sum((S-mu)**g1*log(S-mu))/sum((S-mu)**g1); g2=g1*1.01d0
        dum2=nPnt/g2+sum(log(S-mu))-nPnt*sum((S-mu)**g2*log(S-mu))/sum((S-mu)**g2)
        df=(dum2-dum1)/(g1*.01)
        write(*,'(g0,3(1x,g0),3i)')g1,dum1,dum2,df
        g1=g1-dum1/df !Newton-R's method
        if(isnan(dum1))then
          write(*,*)g1,dum1,dum2 
        end if    
        if(abs(dum1)<1d-10)exit
   end do
   X(1)=g1
   X(2)=(sum((S-mu)**g1)/nPnt)**(1d0/g1)
   END subroutine WeibullMLE4GB
   
   subroutine WeibullMME_Skewness4B(Skew,Scale)
   implicit none
   real*8,intent(in)::Skew! Moment Equation Skewness: m_x^{(3)}/sqrt(m_x^{(2)})^2 Weibull par.β:scale(Characteristic Life),γ:shape(Weibull Slope),μ:shift/location(Waiting Time)
   real*8,intent(out)::Scale !β: Characteristic Life
   real*8::dum1,dum2,g1,g2,df,mu
   integer::nPnt
   g1=1d0; if(Scale>0d0.and.Scale<10)g1=Scale
   dum1=1d0;
   do while(abs(dum1)>1d-10)
        dum1=Skew-(Gamma(1d0+3d0/g1)-3d0*Gamma(1d0+2d0/g1)*Gamma(1d0+1d0/g1)+2d0*Gamma(1d0+1d0/g1)**3)/(Gamma(1d0+2d0/g1)-Gamma(1d0+1d0/g1)**2); g2=g1*1.01d0
        dum2=Skew-(Gamma(1d0+3d0/g2)-3d0*Gamma(1d0+2d0/g2)*Gamma(1d0+1d0/g2)+2d0*Gamma(1d0+1d0/g2)**3)/(Gamma(1d0+2d0/g2)-Gamma(1d0+1d0/g2)**2);
        df=(dum2-dum1)/(g1*.01)
        write(*,'(g0,3(1x,g0),3i)')g1,dum1,dum2,df
        g1=g1-dum1/df !Newton-R's method
        if(isnan(dum1))then
          write(*,*)g1,dum1,dum2 
        end if    
        if(abs(dum1)<1d-10)exit
   end do
   Scale=g1
   END subroutine WeibullMME_Skewness4B
   
!=======================================================================
  SUBROUTINE FITDIS(XX,PP,P_ext,PLKey,isUnique,ErrorKey,iWeight,Xex,LOCD,SCAD,SHAD,fName,xTitle,yTitle,xlBuk,ArkNam)
!=======================================================================
!     NOTATIONS :  Routine  Type:     S  = Subroutine
!                                     DF = Double Precision Function
!                  Variable Type:     I  = Input variable
!                                     O  = Output variable
!                                     L  = Local variable
!                  Variable Decl:     I  = Integer
!                                     D  = Double Precision
!                                     C  = Character
!                                     L  = Logical
!     ROUTINES  :
!
!     Name             Type   Description
!     -----------------------------------------------------------------
!
!     VARIABLES :
!
!     Name      Type   Decl   Description
!     -----------------------------------------------------------------
!     IDIS        I      I    Distribution No.
!     XX(:)       I      D    X-values
!     PP(:)       I      D    P-Values
!     LOCD        IO     D    Location parameter
!     SCAD        IO     D    Scaling  parameter
!     SHAD        IO     D    Shape    parameter
!PROB P_ext       I      D    Probability of extreme value
!IPROBIpStart     O      I    ID for first point in tail curve within the PLKey
!CODE1PLKey       I      D    1) Switch for tail fitting if >min(pp); 2) ProbL threshold if <0 or 0=<PLKey<=min(pp)
!CODE2isUnique    I      L    Code for delete or not delete equal prob. points
!CODE3ErrorKey    I      I    Error calculation
!CODE4iWeight     I      I    Weight factor for fitting (0=equal weight,1=more weight at tail)
!fName            I      C    The file name for plot
!xTitle           I      C    The title for horizontal (x-) axis of the plot
!yTitle           I      C    The title for vertical (y-) axis of the plot
!xlBuk            I      C    The XLS book handle
!AkrNam           I      C    The XLS sheet name contained in XlBuk

!=======================================================================
!

!-----------------------------------------------------------------------
!     Declaration of variables
!-----------------------------------------------------------------------
      use libxl
!     use util_xk,only:reallocate,sort,assert_eq,LSqErr
      IMPLICIT NONE
      real*8,allocatable::XX(:),PP(:)
      real*8,intent(in)::P_ext,PLKey
      integer,intent(in)::isUnique,ErrorKey,iWeight
      real*8,intent(out),optional::Xex(2)
      real*8,intent(inout),optional::LOCD,SCAD,SHAD
      character*(*),intent(in),optional::fName,xTitle,yTitle,ArkNam
      integer,intent(in),optional::xlBuk
      !type(BookHandle),intent(in),optional::xlBuk
      INTEGER::Nxx,Nx,Nk,I,J,IpStart,IW,NEFF,JJ,IU,iuplt,ir,ipo=0,ine=0,jmp=0,mit=0,KSHA
      real*8::LOC,SCA,SHA,ERR,SX,SY,SXX,SXY,SYY,PD,PT,RMP,PROBM,PROBL,ERRD,YDD,YTT,XD,XT,YY,X1,X2,Y1,Y2,XN,YN,dum1=1,dum2=2,dum3=3,MinErr,SHA0
      type(BookHandle)::hxlBuk
      type(SheetHandle)::Ark0,Ark
!-----------------------------------------------------------------------
      real*8,allocatable::X(:),P(:),Neg1mLogP(:),Y(:),YD(:),ErrSum(:)
      integer,allocatable::iPos(:),iWei(:)
      logical::LUnique,Pres_LOCD,Pres_SCAD,Pres_SHAD,Pres_fName,Pres_xTitle,Pres_yTitle,Pres_xlBuk,Pres_ArkNam,isPlot,SolGiven
      character(:),allocatable::VarName
!-----------------------------------------------------------------------
!     Settings or Defaults
!-----------------------------------------------------------------------
      IU=1122; iuplt=1199
!iWeight      ! 1=more weighting towards tail, 0=equal weighting
      
      Pres_fName=Present(fName); Pres_xTitle=Present(xTitle); Pres_yTitle=Present(yTitle); Pres_xlBuk=Present(xlBuk); Pres_ArkNam=Present(ArkNam); 
      Pres_LOCD=Present(LOCD); Pres_SCAD=Present(SCAD); Pres_SHAD=Present(SHAD); 
      if(Pres_yTitle)then; VarName=yTitle; else; VarName=''; endif
      SolGiven=Pres_LOCD.and.Pres_SCAD.and.Pres_SHAD; if(SolGiven)SolGiven=LOCD/=0.and.SCAD>0.and.SHAD>0
      isPlot=Pres_fName.and.Pres_xTitle.and.Pres_yTitle.and.Pres_ArkNam; 
      if(isPlot)isPlot=len_trim(fName)>0.and.len_trim(xTitle)>0.and.len_trim(yTitle)>0.and.len_trim(ArkNam)>0
      if(Pres_xlBuk)hxlBuk%point=xlBuk
!-----------------------------------------------------------------------
!     Copy input tables to local tables
!-----------------------------------------------------------------------
      IF(PLKey<0.0)THEN
        RMP   = -PLKey
        PROBM = 1.0-(1.0-P_ext)/RMP
        PROBL = 1.0-MIN((1.0-P_ext)*RMP*0.5,1.0)
        PROBL = MIN(PROBL,0.87)
      ELSE
        PROBM = 1.0; IF(PLKey>1)return
        PROBL = PLKey
      ENDIF

!!===============================================================================================================================!!
!   1. Remove non-applicable values from numerical probability distribution if data source is not reliable/not well-processed
!   2. Remove equal probability values from numerical distribution
!   3. Remove values outside applicable probability range from numerical distribution
!!===============================================================================================================================!!
      Nxx=ASSERT_EQ(size(XX),size(PP),'Input to FITDIS(XX,PP... has size error!')
      LUnique=merge(.true.,.false.,isUnique==1)
      allocate(iPos(Nxx)); iPos=[(i,i=1,Nxx)]; NK=1
      call sort(XX,iPos,NK,.true.) ! sort the response from smallest to largest and make unique
      PP=PP(iPos); Nxx=NK; call reallocate(XX,Nxx); call reallocate(PP,Nxx);
      iPos=[(i,i=1,Nxx)]; NK=1
      call sort(PP,iPos,NK,LUnique)! sort the response from smallest to largest and make unique
      if(LUnique)then
       XX=XX(iPos); Nxx=Nk; call reallocate(XX,NK); call reallocate(PP,NK);     
      end if
      
      do i=1,5
       Nx=count(PP>ProbL.and.PP<=ProbM); IF(NX==0)return; call reallocate(X,Nx); call reallocate(P,Nx); call reallocate(iWei,Nx)
       IpStart=Nxx-Nx+1; X=XX(IpStart:); P=PP(IpStart:)
       if(Nx<=2.and.Ipstart>1)then; PROBL=PP(IpStart-1)*0.99; else; exit; end if
      end do; call reallocate(Neg1mLogP,Nx); Neg1mLogP=-LOG(1d0-P)

!-----------------------------------------------------------------------
!     Fit tail distribution
!-----------------------------------------------------------------------
    allocate(Y(Nx),ErrSum(Nx)); iWei=[(merge(i,1,iWeight==1),i=1,Nx)]; NEFF=sum(iWei)
    dum3=1d0; ipo=0; ine=0; jmp=0; mit=0
      IF(NX>=2)THEN
        if(Pres_SHAD.and.SHAD>0)then
          SHA=SHAD; 
        else
!1) First screening step to find the domain appearing a local minimum (might be non-unique).
         MinErr=1d10
         do KSHA=1,201                   ! Shape from 0.11 to 10.11
          SHA=0.11d0+(KSHA-1)*0.05d0
!         Establish Scale and Location parameter for selected Shape parameter
          Y=(Neg1mLogP)**(1d0/SHA)  
          call XYSumm(iWei,X,Y,SX,SY,SXX,SXY,SYY)
          dum1=LSqErr(Nx,iWei,X,Neg1mLogP,SHA)!Neg1mLogP=>(-LOG(1d0-P))** 
          if(MinErr>dum1)then !Record the minimum Error
             MinErr=dum1; SHA0=SHA
          end if
         end do; MinErr=1d10; SHAD=SHA0-0.05d0
!2) Second screening step to find the minimum neighborhood within the first domain.
         do KSHA=1,2001
          SHA=SHAD+(KSHA-1)*0.00005d0
!         Establish Scale and Location parameter for selected Shape parameter
          Y=(Neg1mLogP)**(1d0/SHA)  
          call XYSumm(iWei,X,Y,SX,SY,SXX,SXY,SYY)
          dum1=LSqErr(Nx,iWei,X,Neg1mLogP,SHA)!Neg1mLogP=>(-LOG(1d0-P))** 
          if(MinErr>dum1)then !Record the minimum Error
             MinErr=dum1; SHA0=SHA
          end if
         end do
           SHAD=SHA0
           Y=(Neg1mLogP)**(1d0/SHA)
           call XYSumm(iWei,X,Y,SX,SY,SXX,SXY,SYY)
           SCAD=(NEFF*Sxy-Sx*Sy)/(NEFF*Syy-Sy*Sy)
           LOCD=(Sx-SCAD*Sy)/NEFF
           YTT=-LOG(1.0-P_ext)
           XT=SCAD*(YTT**(1/SHAD))+LOCD
           ERRD=MinErr
        End if
!3) Third refining step with a good start and apply modified Newton's method.
        MinErr=1d10; SHA=SHA0
        DO ! SHA=0.11d0+(KSHA-1)*0.01d0
!         Establish Scale and Location parameter for selected Shape parameter
          Y=(Neg1mLogP)**(1d0/SHA)  
          call XYSumm(iWei,X,Y,SX,SY,SXX,SXY,SYY)
          dum1=LSqErr(Nx,iWei,X,Neg1mLogP,SHA) 
          if(MinErr>dum1)then
             MinErr=dum1; SHA0=SHA
          end if  
          dum2=LSqErr(Nx,iWei,X,Neg1mLogP,SHA*1.005d0)!+LSqErr(Nx,iWei,X,Neg1mLogP,SHA*.99d0))
          if(dum2>dum1)then
           if(ine>1)then; ine=0; jmp=jmp+1; else; jmp=0; endif
           ipo=ipo+1
          else
           if(ipo>1)then; ipo=0; jmp=jmp+1; else; jmp=0;  endif
           ine=ine+1
          end if
!          write(*,'(g0,3(1x,g0),3i)')SHA,dum1,dum2,(dum2-dum1)/(SHA*.01),ipo,ine,jmp! 0.718480731860414   79705.2541164183   79705.2541164183 0.10127E-08
          SHA=SHA-0.1*(dum2-dum1)/dum1*(dum3/(jmp+1d0)*(ipo+ine)) !Gradient scaled and can be improved to a function of numbers of downsliding and uprising in case some hard situation.
          if(SHA<0)then
           SHA=SHA+2.01*abs((dum2-dum1)/dum1*(dum3/(jmp+1d0)*(ipo+ine))) !Gradient scaled and can be improved to a function of numbers of downsliding and uprising in case some hard situation.
          elseif(dum1/=dum1.or.dum2/=dum2)then
           SHA=SHA0*1.01; cycle
          end if
!Record the minimum Error
          if(mit>200.and.MinErr<dum1)then
           SHA=SHA0
           Y=(Neg1mLogP)**(1d0/SHA)
           call XYSumm(iWei,X,Y,SX,SY,SXX,SXY,SYY)
           SCA=(NEFF*Sxy-Sx*Sy)/(NEFF*Syy-Sy*Sy)
           LOC=(Sx-SCA*Sy)/NEFF
           ERR=MinErr; exit
          elseif(abs(dum1-dum2)<1d-10)then
           SCA=(NEFF*Sxy-Sx*Sy)/(NEFF*Syy-Sy*Sy)
           LOC=(Sx-SCA*Sy)/NEFF
           ERR=dum1
           exit
          !elseif(mit>100.and.abs(dum1-dum2)<1)then
          ! SHA=SHA0
          ! SCA=(NEFF*Sxy-Sx*Sy)/(NEFF*Syy-Sy*Sy)
          ! LOC=(Sx-SCA*Sy)/NEFF
          ! ERR=dum1
          ! exit
          end if
          if(abs(dum1-dum2)>10.)then
           SCA=(NEFF*Sxy-Sx*Sy)/(NEFF*Syy-Sy*Sy)
           LOC=(Sx-SCA*Sy)/NEFF
           write(*,'(a,3(1x,g17.10))')VarName//'_SHA,SCA,LOC=',SHA,SCA,LOC
          end if    
          if(dum1/=dum1)then; 
             write(*,'(a,3(1x,g17.10))')VarName//'SHA,SCA,LOC=',SHA,SCA,LOC
             write(*,*)'MinErr,ERR=',Err,MinErr,'WEIBULL Fitting failed to converge!...'; pause; 
          end if
          mit=mit+1;
        END DO
      END IF
!-----------------------------------------------------------------------
!     Startimes full fitted distribution
!-----------------------------------------------------------------------
    IF(SolGiven)THEN
      call reallocate(YD,Nx)
       YD=(-LOG(1.0-P))**(1d0/SHAD)
          IF (ErrorKey==1) THEN
             where((X-LOCD)/SCAD>=0d0.or.SHAD>=1d0)ErrSum=iWei*(P-(1d0-EXP(-((X-LOCD)/SCAD)**SHAD)))**2
             ERRD=sum(ErrSum)
           ELSEIF(ErrorKey==3) THEN
             ERRD=sum(iWei*(YD-((X-LOCD)/SCAD))**2)
           ELSE
             ERRD=sum(iWei*(X-SCAD*YD-LOCD)**2)
          ENDIF
      !WRITE(114,'(a,14x,a)')'Distribution','i     w    X(i)        P(i)             Y(i)   PT(i)        YT(i)    PD(i)        YD(i)'
!-----------------------------------------------------------------------
!     If startimes fitted distribution better than tailfit then use startimes fitted distribution
!-----------------------------------------------------------------------
    ELSE
      !WRITE(114,'(a,14x,a)')'Distribution','i     w    X(i)        P(i)             Y(i)   PT(i)        YT(i)'
    END IF
!-----------------------------------------------------------------------
!     Print to temporary file
!-----------------------------------------------------------------------
    IF(isPlot)THEN
      open(iuplt,file=fName)
      write(iuplt,'(a)')'VARIABLES="'//trim(xTitle)//'" "'//yTitle//'", Zone T="Numerical Distribution"'
!     Count number of points with probability below threshold value
      if(hxlBuk%point/=0)then
       Ark0=xlBookGetSheetbyName(hxlBuk,'T1'C);
       if(pres_ArkNam)then
         Ark=xlBookGetSheetbyName(hxlBuk,trim(ArkNam)//''C); if(Ark%point==0)Ark=xlBookAddSheet(hxlBuk,trim(ArkNam)//''C,Ark0)
       else
         Ark=xlBookAddSheet(hxlBuk,"EVFIT"C)
       end if  
       ir=0; call WriteCell(Ark,ir,0,['Distribution','i','w','X(i)','P(i)','Y(i)','PT(i)','YT(i)']); ir=ir+1
      end if
    END IF
!     Print distribution (probability below threshold value)
      DO I=1,Nxx
        IF(PP(I)<=PROBL) THEN
          IF (XX(I)<=LOC) THEN
            PT=0.0
          ELSE
            PT=1d0-EXP(-((XX(I)-LOC)/SCA)**SHA)
          ENDIF
          YTT = -LOG(1d0-PT)         ! Weibull Tail Fitted Distribution
          YY  = -LOG(1d0-PP(I))      ! Numerical Distribution
         IF(SolGiven)THEN
          IF(XX(I)<=LOCD) THEN
            PD=0.0
          ELSE
            PD=1d0-EXP(-((XX(I)-LOCD)/SCAD)**SHAD)
          ENDIF
          YDD = -LOG(1d0-PD)         ! Startimes/Usergiven Distribution
          !WRITE(114,'(a,2(1X,I5),1X,F12.4,3X,E10.4,5X,F8.4,2(1X,E10.4,1X,F8.4))')'Lower distribution : ',-I,0,XX(I),PP(I),YY,PT,YTT,PD,YDD
         ELSEif(Pres_xlBuk)then 
          !WRITE(114,'(a,2(1X,I5),1X,F12.4,3X,E10.4,5X,F8.4,1X,E10.4,1X,F8.4)')'Lower distribution : ',-I,0,XX(I),PP(I),YY,PT,YTT
           if(hxlBuk%point/=0)then
            call WriteCell(Ark,ir,0,'Lower distribution'); call WriteCell(Ark,ir,1,[1d0,dble(-I),0d0,XX(I),PP(I),YY,PT,YTT])
            ir=ir+1
           end if
         END IF 
        ENDIF
       IF(isPlot)write(iuplt,'(f13.3,f8.4)')XX(I),-LOG(1.0-PP(I)) !YY 
      END DO

!     Calculate estimated extreme based on numerical distribution
      XN = 0.0; YN = 0.0
      DO I=1,NX-1
        IF((P_ext>=P(I)).AND.(P_ext<=P(I+1))) THEN
          Y1 = -LOG(1d0-P(I))
          Y2 = -LOG(1d0-P(I+1))
          YN = -LOG(1d0-P_ext)
          X1 = X(I)
          X2 = X(I+1)
          XN = X1+(YN-Y1)*(X2-X1)/(Y2-Y1)
          IF(isPlot)write(iuplt,'(f13.3,f8.4)')XN,YN 
          exit
        ENDIF
      END DO
      
!     Print distribution (probability above threshold value, tailpart)
    IF(isPlot)write(iuplt,'(a)')'Zone T="Weibull Tail Fitted Distribution"'
      DO I=1,NX
	    if((X(I)-LOC)/SCA<0d0.and.SHA<1d0)cycle
        PT=1d0-EXP(-((X(I)-LOC)/SCA)**SHA)
        YTT = -LOG(1d0-PT)   ! Weibull Tail Fitted Distribution
        YY  = -LOG(1d0-P(I)) ! Numerical Distribution
!WRITE(UNIT=114,FMT=1020) IDIS,I,IW,X(I),P(I),YY,PD,YDD,PT,YTT
       IF(SolGiven)then
        PD=1.0-EXP(-((X(I)-LOCD)/SCAD)**SHAD)
        YDD = -LOG(1d0-PD)   ! Startimes/Usergiven Distribution
         !WRITE(114,'(a,2(1X,I5),1X,F12.4,3X,E10.4,5X,F8.4,1X,E10.4,1X,F8.4,3x,E10.4,1X,F8.4)')'Tail distribution  : ',I,iWei(i),X(I),P(I),YY,PT,YTT,PD,YDD
       ELSE
         !WRITE(114,'(a,2(1X,I5),1X,F12.4,3X,E10.4,5X,F8.4,1X,E10.4,1X,F8.4)')'Tail distribution  : ',I,iWei(i),X(I),P(I),YY,PT,YTT
       END IF
       IF(isPlot)THEN
         write(iuplt,'(f13.3,f8.4)')X(I),YTT 
           if(hxlBuk%point/=0)then
            call WriteCell(Ark,ir,0,'Tail distribution'); call WriteCell(Ark,ir,1,[dble(I),dble(iWei(i)),X(I),P(I),YY,PT,YTT])
            ir=ir+1
           end if
       END IF       
      END DO
      YTT = -LOG(1.0-P_ext)
      if(ERRD<MinErr)then
       XT=SCAD*(YTT**(1/SHAD))+LOCD
      else
       XT=SCA*(YTT**(1/SHA))+LOC
      end if 
      Xex=[XT,YTT]
!     Print estimated extremes at specified probability
     IF(SolGiven)then
      YDD = YTT
      XD  = SCAD*(YDD**(1/SHAD))+LOCD
      !WRITE(114,"(a,2(2X,'  - '),1X,F12.4,3X,' -------- ',5X,'   ---- ',1X,E10.4,1X,F8.4,4x,'(',F12.4')')")'TailFit extreme(XD): ',XT,P_ext,YTT,XD
     ELSE 
      !WRITE(114,"(a,2(2X,'  - '),1X,F12.4,3X,' -------- ',5X,'   ---- ',1X,E10.4,1X,F8.4)")'TailFit   extreme  : ',XT,P_ext,YTT
     END IF 
    IF(isPlot)THEN
     write(iuplt,'(f13.3,f8.4)')XT,YTT 
     write(iuplt,'(a)')'Zone T="Estimated 100yr RP Value"' !Repeat the last Extrapolated/Estimated Extreme Valued into a separate Zone
     write(iuplt,'(f13.3,f8.4)')XT,YTT
           if(hxlBuk%point/=0)then
            call WriteCell(Ark,ir,0,'TailFit extreme'); call WriteCell(Ark,ir,1,[0d0,0d0,XT,0d0,0d0,P_ext,YTT])
            ir=ir+1
           end if
    END IF       
!     Print weibull parameters and obtained error levels
      !WRITE(114,'(a)')'AUXILIARY:'
      !WRITE(114,"('ProbL P_ext ProbM  : '2(1X,I5),3(1X,f13.5))")-99,-99,PROBL,P_ext,PROBM
     IF(SolGiven)then
      !WRITE(114,"('TF: loc,sca,sha,err: '2(1X,I5),3(1X,F13.5),1(1X,E12.5),1x,'(ErrD=',E12.5,')')")-99,-99,LOC,SCA,SHA,ERR,ERRD
     ELSE 
      !WRITE(114,"('TF: loc,sca,sha,err: '2(1X,I5),3(1X,F13.5),1(1X,E12.5))")-99,-99,LOC,SCA,SHA,ERR
     END IF 
     if(hxlBuk%point/=0)then
      call WriteCell(Ark,ir,0,'AUXILIARY'); ir=ir+1;
      call WriteCell(Ark,ir,0,'ProbL P_ext ProbM'); call WriteCell(Ark,ir,1,[-99d0,-99d0,PROBL,P_ext,PROBM]); ir=ir+1
      call WriteCell(Ark,ir,0,'TF: loc,sca,sha,err'); call WriteCell(Ark,ir,1,[-99d0,-99d0,LOC,SCA,SHA,ERR])
     end if

    IF(ERRD>MinErr)then
     LOCD = LOC
     SCAD = SCA
     SHAD = SHA
    END IF
  END SUBROUTINE FITDIS

End Module StatPack