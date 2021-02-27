clear
clc
%% PROCESSING
PiyasaTakasTotal = xlsread('PTF-03092018-17092019.xls');
ToplamTuketim=xlsread('GercekZamanliTuketim-03092018-17092019.xls');
YukTahminPlani = xlsread('YukTahminPlani-03092018-16092019.xls');

% dataDates = createDates();% Date predictors %'dd-mm-yyyy''03-Sep-2018'
% save 'dataDates';
% data.Hour = createHours();
% dayOfWeek1 = weekday(dataDates);
% save 'dayOfWeek1';
% dayOfWeek = createDayOfWeek();
% save 'dayOfWeek';
% data.DogalgazSTPFiyatGRF1 = createNG()
% isWorkingDay = createHoliday(); % Non-business days
% HourxWeek = data.Hour.*dayOfWeek;

NGPrices=xlsread('Gaz_Referans_Fiyati_(GRF)_2018-09-03_2019-09-17.xls');
kk=1;
for i=1:1:length(NGPrices(:,2))
    NGPricesnew(kk:kk+23,1)=NGPrices(i,2);
    kk=kk+24;
end

Haftasonusilinik22=xlsread('Haftasonusilinik22.xls');

CoalPrice =  xlsread('CoalPrice.xls');
kk=1;
for i=1:1:length(CoalPrice)
    CoalPricesnew(kk:kk+23,1)=CoalPrice(i,1);
    kk=kk+24;
end
OilPrices = xlsread('OilPrices.xls');
kk=1;
for i=1:1:length(OilPrices(:,2))
    OilPricesnew(kk:kk+23,1)=OilPrices(i,2);
    kk=kk+24;
end
%% TRAINING
% Lagged Load inputs
% prevDaySameHourLoad = [NaN(24,1); ToplamTuketim(1:end-24)];
% prevWeekSameHourLoad = [NaN(168,1); ToplamTuketim(1:end-168)];
% prev24HrAveLoad = filter(ones(1,24)/24, 1, ToplamTuketim);
%
%
% prevDaySameHourYukTahminPlani = [NaN(24,1); YukTahminPlani(1:end-24)];
% prevWeekSameHourYukTahminPlani = [NaN(168,1); YukTahminPlani(1:end-168)];
% prev24HrAveYukTahminPlani = filter(ones(1,24)/24, 1, YukTahminPlani);
%
% % Lagged electricity price inputs
% prevDaySameHourPrice = [NaN(24,1); PiyasaTakasTotal(1:end-24)];
% prevWeekSameHourPrice = [NaN(168,1); PiyasaTakasTotal(1:end-168)];
% prev24HrAvePrice = filter(ones(1,24)/24, 1, PiyasaTakasTotal);
%
% % Lagged NG price inputs
% prevDayNGPrice = [NaN(24,1); NGPricesnew(1:end-24)];
% prevWeekNGPrice = [NaN(168,1); NGPricesnew(1:end-168)];
% prev24AveNGPrice = filter(ones(1,24)/24, 1, NGPricesnew);
%
% % Lagged euro inputs
% prevDayEuroParity = [NaN(24,1); Haftasonusilinik22(1:end-24)];
% prevWeekEuroParity = [NaN(168,1); Haftasonusilinik22(1:end-168)];
% prev24AveEuroParity = filter(ones(1,24)/24, 1, Haftasonusilinik22);
%
% % Lagged Oil inputs
% prevDayOilPrice = [NaN(24,1); OilPricesnew(1:end-24)];
% prevWeekOilPrice = [NaN(168,1); OilPricesnew(1:end-168)];
% prev24AveOilPrice = filter(ones(1,24)/24, 1, OilPricesnew);
%
% % Lagged Coal inputs
% prevDayCoalPrice = [NaN(24,1); CoalPricesnew(1:end-24)];
% prevWeekCoalPrice = [NaN(168,1); CoalPricesnew(1:end-168)];
% prev24AveCoalPrice = filter(ones(1,24)/24, 1, CoalPricesnew);

% X = [ones(8760,1) data.Hour dayOfWeek...
%     isWorkingDay HourxWeek data.GercekZamanliTuketim...
%     prevDaySameHourLoad prevWeekSameHourLoad prev24HrAveLoad...
%     prevDaySameHourPrice prevWeekSameHourPrice prev24HrAvePrice...
%     prevDayNGPrice prevWeekNGPrice prev24AveNGPrice...
%     prevDayOilPrice prevWeekOilPrice prev24AveOilPrice...
%     prevDayCoalPrice prevWeekCoalPrice prev24AveCoalPrice...
%     prevDayEuroParity prevWeekEuroParity prev24AveEuroParity];
k=1;
j=0;
for i=24:24:336
    PiyasaTakasTotal1(:,k)=PiyasaTakasTotal(8761+j:8760+i,1);
    k=k+1;
    j=j+24;
end
%% PRICE OUTLIER
X1newSeptember = xlsread('X1newSeptember.xls');
X1new = xlsread('X1new');

satir_baslangic=1;
sutun=[1:9 10:21];
%% VALIDATION
j=24; rr=1; l=0; rrr=1; rrrr=1;
for i = 0:24:312
    mdl = fitlm(X1new(satir_baslangic:8760+i,sutun), PiyasaTakasTotal(satir_baslangic:8760+i));
    [Ypred,YCI] = predict(mdl, X1newSeptember((i+1):(i+24),sutun),'Alpha',0.01,'Simultaneous',true);
%     [betas,bint,r,rint,stats]= regress(PiyasaTakasTotal(satir_baslangic:8760+i),X1new(satir_baslangic:8760+i,sutun), 0.05);
%     bintlowermatrices(:) = bint(:,1);
%     bintuppermatrices(:) = bint(:,2);
%     betasmatrices(:,rr)=betas(:);
%     ValidationSeptember(:,rr)=X1newSeptember((i+1):(i+24),sutun)*betas;
    ValidationSeptember(:,rr) = Ypred;
    LowerBounds(:,rr) = YCI(:,1);
    UpperBounds(:,rr) = YCI(:,2);
    RealSeptember(:,rr) = PiyasaTakasTotal1(:,rr);
    error=0;
    for k = 1:1:24
        error=error+abs((RealSeptember(k,rr)-ValidationSeptember(k,rr))/RealSeptember(k,rr));
    end
    errorr(rr)=(error/24)*100;
    %
    error=0;
    for k = 1:1:24
        error=error+(RealSeptember(k,rrr)-ValidationSeptember(k,rrr))^2;
    end
    errorRMSE(rrr)=sqrt(error/24);
    %
    error=0;
    for k = 1:1:24
        error=error+(ValidationSeptember(k,rrrr)-RealSeptember(k,rrrr));
    end
    errorMSD(rrrr)=error/24;
    %
    Validation14days(l+1:l+24,1)=ValidationSeptember(:,rr);
    LowerBounds14days(l+1:l+24,1)=LowerBounds(:,rr);
    UpperBounds14days(l+1:l+24,1)=UpperBounds(:,rr);

    l=l+24;
    rr=rr+1;
    rrr=rrr+1;
    rrrr=rrrr+1;
end
err = [LowerBounds14days; UpperBounds14days];
%% PLOT
% oneyearandmonths(12); %1-14 arasında parantezin içine sayı yazılabilir
% print(gcf,'RollingForecastMethodsErrorGraphs.png','-dpng','-r800');
% 
% % ngoilcoaleurograph(12); %1-14 arasında parantezin içine sayı yazılabilir
% % laggedpricesvslaggednonprices(12); %1-14 arasında parantezin içine sayı yazılabilir
% scatterchartpricevsnonprices(12); %1-14 arasında parantezin içine sayı yazılabilir
% scatterngoilcoaleurograph(12);
% scatteroneyearandmonths(12);
%% FORECASTED FORECAST PRICE
j=24; rr=1; t=1; e=1; x=1; y=1; l=0; rrr=1; rrrr=1;
for i = 0:24:312
    if i >=24
        X1newSeptember(i+1:i+24,10) = ForecastedValidationSeptember(:,e);
        v1 = filter(ones(1,24)/24, 1, X1new(1:8760+i,10));
        X1newSeptember(i+1:i+24,12) = v1(8760+i-23:8760+i,1);
        e=e+1;
    end
    if i >=48
        X1new(8760+i-23:8760+i,10) = ForecastedValidationSeptember(:,t);
        v2 = filter(ones(1,24)/24, 1, X1new(1:8760+i,10));
        X1new(8760+i-23:8760+i,12) = v2(8760+i-23:8760+i,1);
        t=t+1;
    end
    if i>=168
        X1new(8760+i-23:8760+i,11) = ForecastedValidationSeptember(:,x);
        x=x+1;
    end
    if i>=192
        X1newSeptember((i+1):(i+24),11) = ForecastedValidationSeptember(:,y);
        y=y+1;
    end
    mdl2 = fitlm(X1new(satir_baslangic:8760+i,sutun), PiyasaTakasTotal(satir_baslangic:8760+i));
    [Ypred2,YCI2] = predict(mdl2, X1newSeptember((i+1):(i+24),sutun),'Alpha',0.01,'Simultaneous',true);
%     [betas,bint,r,rint,stats]= regress(PiyasaTakasTotal(satir_baslangic:8760+i),X1new(satir_baslangic:8760+i,sutun), 0.05);
%     ForecastedValidationSeptember(:,rr)=X1newSeptember((i+1):(i+24),sutun)*betas;
    ForecastedValidationSeptember(:,rr) = Ypred2;
    FFLowerBounds(:,rr) = YCI2(:,1);
    FFUpperBounds(:,rr) = YCI2(:,2);
    RealSeptember(:,rr) = PiyasaTakasTotal1(:,rr);
    error=0;
    for k = 1:1:24
        error=error+abs((RealSeptember(k,rr)-ForecastedValidationSeptember(k,rr))/RealSeptember(k,rr));
    end
    Forecasterror(rr)=(error/24)*100;
    %
    error=0;
    for k = 1:1:24
        error=error+(RealSeptember(k,rrr)-ForecastedValidationSeptember(k,rrr))^2;
    end
    ForecasterrorRMSE(rrr)=sqrt(error/24);
    %
    error=0;
    for k = 1:1:24
        error=error+(ForecastedValidationSeptember(k,rrrr)-RealSeptember(k,rrrr));
    end
    ForecasterrorMSD(rrrr)=error/24;
    %
    
    Forecastforecast14days(l+1:l+24,1)=ForecastedValidationSeptember(:,rr);
    FFLowerBounds14days(l+1:l+24,1)=FFLowerBounds(:,rr);
    FFUpperBounds14days(l+1:l+24,1)=FFUpperBounds(:,rr);
    
    l=l+24;
    PiyasaTakasTotal(8761+i:8760+i+24,1)=ForecastedValidationSeptember(:,rr);%Regressiondaki PTF yi değiştiriyor.
    rr=rr+1;
    rrr=rrr+1;
    rrrr=rrrr+1;
end
%% PLOT
% figure
% pp1=plot(errorr)
% hold on
% pp2=plot(Forecasterror)
% pp1.LineWidth = 2;
% pp2.LineWidth = 2;
% pp1.Marker = 'o';
% pp2.Marker = 'o';
% legend({'Regular Forecast Error','Dynamic Forecast Error'},'Location','northwest','NumColumns',2);
% title('Mean Absolute Percentage Error');
% xlabel('Forecasted Day Numbers');
% ylabel('Error rates');
%
figure
x = 1:1:288';
mm=plot(Validation14days(1:288),'b');
hold on;
ll=plot(Forecastforecast14days(1:288),'g');
h2 = fill([x,fliplr(x)],[UpperBounds14days(1:288)',fliplr(LowerBounds14days(1:288)')],'b','facealpha',0.3);
h3 = fill([x,fliplr(x)],[FFUpperBounds14days(1:288)',fliplr(FFLowerBounds14days(1:288)')],'g','facealpha',0.3);
q=reshape(RealSeptember,24*14,1);
nn=plot(q(1:288),'r');
hold off;
legend({'Regular Forecast Price','Dynamic Forecast Price','Confidence Interval (%99)','Confidence Interval (%99)','Real Price'},'FontSize',9,'Location','northwest','NumColumns',3);
mm.LineWidth = 2;nn.LineWidth = 2;ll.LineWidth = 2;
title('Estimated Prices vs Real Price','FontSize',12);
xlabel('Hours','FontSize',12);
ylabel('Price (TL/MWh)','FontSize',12);
print(gcf,'RealPredicted.png','-dpng','-r800');
