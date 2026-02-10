% Practical assignment

clc; clear; close all
%%

% file/ filepath 
filepath = ("Planning_Data_with_NPV_values.xlsx");


%                          DATA: PLAN_8MT

% Set up the Import Options and import the data
opts = spreadsheetImportOptions("NumVariables", 20);

% Specify sheet and range
opts.Sheet = "Plan_8Mt";
opts.DataRange = "B3:U13";

% Specify column names and types
opts.VariableNames = ["Period", "Tonnes", "Mill1", "Mill_Au_GRADEgt", "Waste", "Stockpiletin", "Stockpiletout", "RecoveryRate", "RecoveredGoldoz", "UnitProcessingCostUSDtn", "processingCostUSD", "MiningCostInflation", "MiningCostUSD", "CapitalExpenditureUSD", "GoldPriceUSDoz", "RevenueFromGoldRecoveredUSD", "TaxAndRoyalty", "ProfitAfterTaxUSD", "DicountFactor", "DicountedNPVUSD"];
opts.VariableTypes = ["double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double"];

% Import the data 
% CHANGE THE FILE PATH TO YOUR OWN !!!
plan_8mt = readtable(filepath, opts, "UseExcel", false);

% Convert to output type
plan_8mt = table2array(plan_8mt);

% Clear temporary variables
clear opts

%                         DATA: PLAN_6MT

% Set up the Import Options and import the data
opts = spreadsheetImportOptions("NumVariables", 20);

% Specify sheet and range
opts.Sheet = "Plan_6Mt";
opts.DataRange = "B3:U16";

% Specify column names and types
opts.VariableNames = ["Period", "Tonnes", "Mill1", "Mill_Au_GRADEgt", "Waste", "Stockpiletin", "Stockpiletout", "RecoveryRate", "RecoveredGoldoz", "UnitProcessingCostUSDtn", "processingCostUSD", "MiningCostInflation", "MiningCostUSD", "CapitalExpenditureUSD", "GoldPriceUSDoz", "RevenueFromGoldRecoveredUSD", "TaxAndRoyalty", "ProfitAfterTaxUSD", "DicountFactor", "DicountedNPVUSD"];
opts.VariableTypes = ["double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double"];

% Import the data

plan_6mt = readtable(filepath, opts, "UseExcel", false);

% Convert to output type
plan_6mt = table2array(plan_6mt);

% Clear temporary variables
clear opts

%                          DATA: PLAN_4MT

% Set up the Import Options and import the data
opts = spreadsheetImportOptions("NumVariables", 20);

% Specify sheet and range
opts.Sheet = "Plan_4Mt";
opts.DataRange = "B3:U22";

% Specify column names and types
opts.VariableNames = ["Period", "Tonnes", "Mill1", "Mill_Au_GRADEgt", "Waste", "Stockpiletin", "Stockpiletout", "RecoveryRate", "RecoveredGoldoz", "UnitProcessingCostUSDtn", "processingCostUSD", "MiningCostInflation", "MiningCostUSD", "CapitalExpenditureUSD", "GoldPriceUSDoz", "RevenueFromGoldRecoveredUSD", "TaxAndRoyalty", "ProfitAfterTaxUSD", "DicountFactor", "DicountedNPVUSD"];
opts.VariableTypes = ["double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double"];

% Import the data
% CHANGE THE FILE PATH TO YOUR OWN !!!
plan_4mt = readtable(filepath, opts, "UseExcel", false);

% Convert to output type
plan_4mt = table2array(plan_4mt);

% Clear temporary variables
clear opts



%%                          Calculate gold price data 

% First lets prepare golds price data:

% Import data from excel
data = readtable('DataSheet.xlsx');
clc

% Columns
time = datetime(data{:,1}, 'InputFormat', 'dd-MMM-yyyy');
price = data{:,2}; 
price = str2double(price);
price_change = data{:,3}; 
price_change(isnan(price_change)) = 0;

% Extract month and year
months = month(time); 
years = year(time);   

% Create matrix
gold_matrix = [years, months, price, price_change]; % maby not needed

% Lets analyze gold price trend

% Random seed 
rng(0);

% Parametrit GBM:lle
num_sim = 100; % Count of simulaations
dt = 1; %
t = (1)';

S0 = 1200; % Gols latest price
mu = mean(price_change) / dt; % Drift
sigma = std(price_change) / sqrt(dt); % Volatiliteetti


% GBM-model
gold_gbm = gbm(mu, sigma, "StartState", S0);

% 100 Monte Carlo simulaations
gold_sim = gold_gbm.simBySolution(length(plan_8mt(:,1))-1, 'DeltaTime', dt, 'NTrials', num_sim);
gold_sim = squeeze(gold_sim);

% Plotting
figure;
plot(gold_sim);
xlabel('Years');
ylabel('Gold Price');
title('Gold Price Prediction using GBM');
grid on;


% Make matrix that have different gold prices for performance analysis:
clc
% bottom_5%
idx5 = find(gold_sim(end,:) == max(mink(gold_sim(end,:), 5))); %Find 5th smallest gold price column
gold_price_bottom5 = gold_sim(:, idx5); 

% top 5%
idx95 = find(gold_sim(end,:) == max(mink(gold_sim(end,:), 95))); %Find 5th largest gold price column
gold_price_top5 = gold_sim(:, idx95); 

% median
idx_median = find(gold_sim(end,:) == max(mink(gold_sim(end,:), 50))); 
gold_price_median = gold_sim(:, idx_median); 

% Gold price is fixed 1200
gold_price_fixed = ones(length(gold_price_median), 1)*1200;

% Matrix
gold_price_matrix = [gold_price_fixed, gold_price_bottom5, gold_price_median, gold_price_top5];


%%                                   CALCULATE Plan_8mt performance 

% Define variables

inflation = 0.02; 
mining_cost_initial = 3.10;
discount_rate = 0.10;

Period = plan_8mt(:, 1);
Tonnes = plan_8mt(:, 2);
Mill1 = plan_8mt(:, 3);
Mill_Au_GRADE = plan_8mt(:, 4);
Recovery_rate = plan_8mt(:, 8);
% Recovery_rate(1:end) = Recovery_rate(1:end) + 0.18;
Unit_processing_cost = plan_8mt(:, 10); 
Capital_expenditure = plan_8mt(:, 14);
Gold_price = gold_price_matrix;
Tax_and_royalty = plan_8mt(:, 17);



gold_mine_model = Simulink.SimulationInput('gold_mine');

% Constant
gold_mine_model = gold_mine_model.setVariable('inflation', inflation);
gold_mine_model = gold_mine_model.setVariable('mining_cost_initial', mining_cost_initial);
gold_mine_model = gold_mine_model.setVariable('discount_rate', discount_rate);

% Initialization
Discounted_NPV = zeros(length(Period),1);
Discounted_nvp_8mt_matrix = zeros(length(Period), 4);


for i = 1:length(gold_price_matrix(1,:))


    for j = 1:length(Period)
        
        % Change every round (period)
        gold_mine_model = gold_mine_model.setVariable('Period', Period(j));
        gold_mine_model = gold_mine_model.setVariable('Tonnes', Tonnes(j));
        gold_mine_model = gold_mine_model.setVariable('Mill1', Mill1(j));
        gold_mine_model = gold_mine_model.setVariable('Mill_Au_GRADE', Mill_Au_GRADE(j));
        gold_mine_model = gold_mine_model.setVariable('Recovery_rate', Recovery_rate(j));
        gold_mine_model = gold_mine_model.setVariable('Unit_processing_cost', Unit_processing_cost(j));
        gold_mine_model = gold_mine_model.setVariable('Capital_expenditure', Capital_expenditure(j));
        gold_mine_model = gold_mine_model.setVariable('Gold_price', Gold_price (j,i));
        gold_mine_model = gold_mine_model.setVariable('Tax_and_royalty', Tax_and_royalty(j));
          
        sim_res = sim(gold_mine_model); % Sim model
        
        Discounted_NPV(j,1) = sim_res.yout{1}.Values.Data; % create discounted nvp vector
    
    end

    Discounted_nvp_8mt_matrix(:, i) = Discounted_NPV;

end

% Results from Plan 8mt

Discounted_nvp_8mt_total = sum(Discounted_nvp_8mt_matrix, 1);
disp('Sum of discounted NPVs in all scenarios: ')
disp(num2str(Discounted_nvp_8mt_total))

figure
hold on
gp_fixed = plot(Period, Discounted_nvp_8mt_matrix(:, 1), 'k*-', 'LineWidth', 1);
gp_bottom5 = plot(Period, Discounted_nvp_8mt_matrix(:, 2), 'r*-', 'LineWidth', 1);
gp_median = plot(Period, Discounted_nvp_8mt_matrix(:, 3), 'b*-', 'LineWidth', 1);
gp_top5 = plot(Period, Discounted_nvp_8mt_matrix(:, 4), 'g*-', 'LineWidth', 1);
xlabel('Period')
ylabel('Discounted NPV (USD)')
title('Plan 8mt Discounted NPV vs. Time')
legend([gp_fixed, gp_bottom5, gp_median, gp_top5], 'g-price fixed (1200)', 'g-price bottom 5 %', 'g-price median','g-price top 5 %');
grid on











%%                              PLAN 6MT STARTS


% 100 Monte Carlo simulaations
gold_sim = gold_gbm.simBySolution(length(plan_6mt(:,1))-1, 'DeltaTime', dt, 'NTrials', num_sim);
gold_sim = squeeze(gold_sim);

% Plotting
figure;
plot(gold_sim);
xlabel('Years');
ylabel('Gold Price');
title('Gold Price Prediction using GBM');
grid on;


% Make matrix that have different gold prices for performance analysis:
clc
% bottom_5%
idx5 = find(gold_sim(end,:) == max(mink(gold_sim(end,:), 5))); %Find 5th smallest gold price column
gold_price_bottom5 = gold_sim(:, idx5); 

% top 5%
idx95 = find(gold_sim(end,:) == max(mink(gold_sim(end,:), 95))); %Find 5th largest gold price column
gold_price_top5 = gold_sim(:, idx95); 

% median
idx_median = find(gold_sim(end,:) == max(mink(gold_sim(end,:), 50))); 
gold_price_median = gold_sim(:, idx_median); 

% Gold price is fixed 1200
gold_price_fixed = ones(length(gold_price_median), 1)*1200;

% Matrix
gold_price_matrix = [gold_price_fixed, gold_price_bottom5, gold_price_median, gold_price_top5];

%%                          CALCULATE Plan_6mt performance
% clc
% Define variables

inflation = 0.02; 
mining_cost_initial = 3.10;
discount_rate = 0.10;

Period = plan_6mt(:, 1);
Tonnes = plan_6mt(:, 2);
Mill1 = plan_6mt(:, 3);
Mill_Au_GRADE = plan_6mt(:, 4);
Recovery_rate = plan_6mt(:, 8);
%Recovery_rate([1,2,3]) = Recovery_rate([1,2,3]) + [0.3; 0.25; 0.20];
%Recovery_rate(4:end) = Recovery_rate(4:end) + 0.15;
Unit_processing_cost = plan_6mt(:, 10);
Capital_expenditure = plan_6mt(:, 14);
Gold_price = gold_price_matrix;
Tax_and_royalty = plan_6mt(:, 17);



gold_mine_model = Simulink.SimulationInput('gold_mine');

% Constant
gold_mine_model = gold_mine_model.setVariable('inflation', inflation);
gold_mine_model = gold_mine_model.setVariable('mining_cost_initial', mining_cost_initial);
gold_mine_model = gold_mine_model.setVariable('discount_rate', discount_rate);

% Initialization
Discounted_NPV = zeros(length(Period),1);
Discounted_nvp_6mt_matrix = zeros(length(Period), 4);


for i = 1:length(gold_price_matrix(1,:))


    for j = 1:length(Period)
        
        % Change every round (period)
        gold_mine_model = gold_mine_model.setVariable('Period', Period(j));
        gold_mine_model = gold_mine_model.setVariable('Tonnes', Tonnes(j));
        gold_mine_model = gold_mine_model.setVariable('Mill1', Mill1(j));
        gold_mine_model = gold_mine_model.setVariable('Mill_Au_GRADE', Mill_Au_GRADE(j));
        gold_mine_model = gold_mine_model.setVariable('Recovery_rate', Recovery_rate(j));
        gold_mine_model = gold_mine_model.setVariable('Unit_processing_cost', Unit_processing_cost(j));
        gold_mine_model = gold_mine_model.setVariable('Capital_expenditure', Capital_expenditure(j));
        gold_mine_model = gold_mine_model.setVariable('Gold_price', Gold_price (j,i));
        gold_mine_model = gold_mine_model.setVariable('Tax_and_royalty', Tax_and_royalty(j));
          
        sim_res = sim(gold_mine_model); % Sim model
        
        Discounted_NPV(j,1) = sim_res.yout{1}.Values.Data; % create discounted nvp vector
    
    end

    Discounted_nvp_6mt_matrix(:, i) = Discounted_NPV;

end

% Results from Plan 6mt

Discounted_nvp_6mt_total = sum(Discounted_nvp_6mt_matrix, 1);
disp('Sum of discounted NPVs in all scenarios: ')
disp(num2str(Discounted_nvp_6mt_total))
figure
hold on
gp_fixed = plot(Period, Discounted_nvp_6mt_matrix(:, 1), 'k*-', 'LineWidth', 1);
gp_bottom5 = plot(Period, Discounted_nvp_6mt_matrix(:, 2), 'r*-', 'LineWidth', 1);
gp_median = plot(Period, Discounted_nvp_6mt_matrix(:, 3), 'b*-', 'LineWidth', 1);
gp_top5 = plot(Period, Discounted_nvp_6mt_matrix(:, 4), 'g*-', 'LineWidth', 1);
xlabel('Period')
ylabel('Discounted NPV (USD)')
title('Plan 6mt Discounted NPV vs. Time')
legend([gp_fixed, gp_bottom5, gp_median, gp_top5], 'g-price fixed (1200)', 'g-price bottom 5 %', 'g-price median','g-price top 5 %');
grid on






%%                                  PLAN 4MT STARTS




% 100 Monte Carlo simulaations
gold_sim = gold_gbm.simBySolution(length(plan_4mt(:,1))-1, 'DeltaTime', dt, 'NTrials', num_sim);
gold_sim = squeeze(gold_sim);

% Plotting
figure;
plot(gold_sim);
xlabel('Years');
ylabel('Gold Price');
title('Gold Price Prediction using GBM');
grid on;


% Make matrix that have different gold prices for performance analysis:
clc
% bottom_5%
idx5 = find(gold_sim(end,:) == max(mink(gold_sim(end,:), 5))); %Find 5th smallest gold price column
gold_price_bottom5 = gold_sim(:, idx5); 

% top 5%
idx95 = find(gold_sim(end,:) == max(mink(gold_sim(end,:), 95))); %Find 5th largest gold price column
gold_price_top5 = gold_sim(:, idx95); 

% median
idx_median = find(gold_sim(end,:) == max(mink(gold_sim(end,:), 50))); 
gold_price_median = gold_sim(:, idx_median); 

% Gold price is fixed 1200
gold_price_fixed = ones(length(gold_price_median), 1)*1200;

% Matrix
gold_price_matrix = [gold_price_fixed, gold_price_bottom5, gold_price_median, gold_price_top5];



%%                          CALCULATE Plan_4mt performance
clc
% Define variables

inflation = 0.02; 
mining_cost_initial = 3.10;
discount_rate = 0.10;

Period = plan_4mt(:, 1);
Tonnes = plan_4mt(:, 2);
Mill1 = plan_4mt(:, 3);
Mill_Au_GRADE = plan_4mt(:, 4);
Recovery_rate = plan_4mt(:, 8);
% Recovery_rate([1,2]) = Recovery_rate([1,2]) + [0.25; 0.15];
% Recovery_rate(3:end) = Recovery_rate(3:end) + 0.11;
Unit_processing_cost = plan_4mt(:, 10);
Capital_expenditure = plan_4mt(:, 14);
Gold_price = gold_price_matrix;
Tax_and_royalty = plan_4mt(:, 17);

gold_mine_model = Simulink.SimulationInput('gold_mine');

% Constant
gold_mine_model = gold_mine_model.setVariable('inflation', inflation);
gold_mine_model = gold_mine_model.setVariable('mining_cost_initial', mining_cost_initial);
gold_mine_model = gold_mine_model.setVariable('discount_rate', discount_rate);

% Initialization
Discounted_NPV = zeros(length(Period),1);
Discounted_nvp_4mt_matrix = zeros(length(Period), 4);


for i = 1:length(gold_price_matrix(1,:))


    for j = 1:19 %length(Period)
        
        % Change every round (period)
        gold_mine_model = gold_mine_model.setVariable('Period', Period(j));
        gold_mine_model = gold_mine_model.setVariable('Tonnes', Tonnes(j));
        gold_mine_model = gold_mine_model.setVariable('Mill1', Mill1(j));
        gold_mine_model = gold_mine_model.setVariable('Mill_Au_GRADE', Mill_Au_GRADE(j));
        gold_mine_model = gold_mine_model.setVariable('Recovery_rate', Recovery_rate(j));
        gold_mine_model = gold_mine_model.setVariable('Unit_processing_cost', Unit_processing_cost(j));
        gold_mine_model = gold_mine_model.setVariable('Capital_expenditure', Capital_expenditure(j));
        gold_mine_model = gold_mine_model.setVariable('Gold_price', Gold_price (j,i));
        gold_mine_model = gold_mine_model.setVariable('Tax_and_royalty', Tax_and_royalty(j));
          
        sim_res = sim(gold_mine_model); % Sim model
        
        Discounted_NPV(j,1) = sim_res.yout{1}.Values.Data; % create discounted nvp vector
    
    end

    Discounted_nvp_4mt_matrix(:, i) = Discounted_NPV;

end

% Results from Plan 4mt

Discounted_nvp_4mt_total = sum(Discounted_nvp_4mt_matrix, 1);
disp('Sum of discounted NPVs in all scenarios: ')
disp(num2str(Discounted_nvp_4mt_total))

figure
hold on
gp_fixed = plot(Period, Discounted_nvp_4mt_matrix(:, 1), 'k*-', 'LineWidth', 1);
gp_bottom5 = plot(Period, Discounted_nvp_4mt_matrix(:, 2), 'r*-', 'LineWidth', 1);
gp_median = plot(Period, Discounted_nvp_4mt_matrix(:, 3), 'b*-', 'LineWidth', 1);
gp_top5 = plot(Period, Discounted_nvp_4mt_matrix(:, 4), 'g*-', 'LineWidth', 1);
xlabel('Period')
ylabel('Discounted NPV (USD)')
title('Plan 4mt Discounted NPV vs. Time')
legend([gp_fixed, gp_bottom5, gp_median, gp_top5], 'g-price fixed (1200)', 'g-price bottom 5 %', 'g-price median','g-price top 5 %');
grid on

%% ALL RESULTS

disp(' ')
disp('Sum of discounted NPVs in all scenarios: ')
disp(' ')
disp('Plan 8mt: ')
disp(['g-price fixed1200  ', ' g-price bottom 5% ', '  g-price median ','     g-price top 5% '])
disp(num2str(Discounted_nvp_8mt_total))
disp(' ')
disp('Plan 6mt: ')
disp(['g-price fixed1200  ', ' g-price bottom 5% ', '  g-price median ','     g-price top 5% '])
disp(num2str(Discounted_nvp_6mt_total))
disp(' ')
disp('Plan 4mt: ')
disp(['g-price fixed1200  ', ' g-price bottom 5% ', ' g-price median ','     g-price top 5% '])
disp(num2str(Discounted_nvp_4mt_total))


%%
