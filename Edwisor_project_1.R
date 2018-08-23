#Clear the environment
rm(list = ls())

#Set working Directory
setwd("D:/Project")
getwd()

#Load XLSX library
library("xlsx")

#Read the data

data = read.xlsx("Absenteeism_at_work_Project.xls", sheetIndex = 1)

#observe the class of data object
class(data)

#Observe the dimension of data
dim(data)

#get the names of variable
colnames(data)


#Check the structure of the dataset variables
str(data)

#-- note that all the variables are of numeric type

#-- Now we have to observe which are the variables that are continuous and which are categorical variables.

#-- we will convert the varaibles in their respective data types but first we need to do missing value analysis

#MISSING VALUE ANALYSIS
sum(is.na(data$Transportation.expense))
missing_value = data.frame(apply(data,2,function(x){sum(is.na(x))}))

names(missing_value)[1] = "Missing_data"

missing_value$Variables = row.names(missing_value)

row.names(missing_value) = NULL

#Reaarange the columns
missing_value = missing_value[, c(2,1)]

#place the variables according to their number of missing values.
missing_value = missing_value[order(-missing_value$Missing_data),]

#Calculate the missing value percentage
missing_value$percentage = (missing_value$Missing_data/nrow(data) )* 100

#Store the missing value information in a csv file
write.csv(missing_value,"Project_Missing_value.csv", row.names = F)

#--> Since the calculated missing percentage is less than 5% 
#--> Thus there is no need to drop any variable due to less number of data

#Now generate a missing value in the dataset and try various imputation methods on it

#make a backup for data object
#data_back = data

## create a missing value in any variable
#data$Transportation.expense[57] #225
#data$Transportation.expense[57] = NA

#ACTUTAL vALUE 92
#Median 95
#KNN   94.22 
#Mean 94.59


##Now apply mean method to impute this generated value and observe the result
#data$Transportation.expense[is.na(data$Transportation.expense)]  =  mean(data$Transportation.expense, na.rm = T)# 221.03

##Now apply median method to impute the missing value
#data$Transportation.expense[is.na(data$Transportation.expense)] = median(data$Transportation.expense, na.rm = T)#225

#Now convert the variables into their respective types

data$ID = as.factor(as.character(data$ID))
data$Day.of.the.week = as.factor(as.character(data$Day.of.the.week))
data$Education = as.factor(as.character(data$Education))
data$Social.drinker = as.factor(as.character(data$Social.drinker))
data$Social.smoker = as.factor(as.character(data$Social.smoker))
data$Reason.for.absence = as.factor(as.character(data$Reason.for.absence))
data$Seasons = as.factor(as.character(data$Seasons))
data$Month.of.absence = as.factor(as.character(data$Month.of.absence))
data$Disciplinary.failure = as.factor(as.character(data$Disciplinary.failure))


#now apply KNN method to impute

library(DMwR)
data = knnImputation(data,k=3)

#-- Both median and KNN are giving same result we can choose any one for furthur implmentation
#-- On analysis of different data using median and KNN i find that KNN is more accurate than median for imputation

#Now load the data again and impute the missing value by KNN
#Check presence of missing values once to confirm

apply(data,2, function(x){sum(is.na(x))})

#NO missing value is found

 
#create subset of the dataset which have only numeric varaiables

numeric_index = sapply(data, is.numeric)
numeric_data = data[,numeric_index]
numeric_data = as.data.frame(numeric_data)

n_data = colnames(numeric_data)[-12]

class(n_data)


#draw the boxplot to detect outliers


library("ggplot2")



for (i in 1:length(n_data)) {
  assign(paste0("gn",i), ggplot(aes_string( y = (n_data[i]), x= "Absenteeism.time.in.hours") , data = subset(data)) + 
    stat_boxplot(geom = "errorbar" , width = 0.5) +
    geom_boxplot(outlier.color = "red", fill = "grey", outlier.shape = 20, outlier.size = 1, notch = FALSE)+
    theme(legend.position = "bottom")+
    labs(y = n_data[i], x= "Absenteeism.time.in.hours")+
    ggtitle(paste("Boxplot" , n_data[i])))
    #print(i)
}

options(warn = -1)

#Now plotting the plots

gridExtra::grid.arrange(gn1, gn2,gn3, ncol=3)
gridExtra::grid.arrange(gn4,gn5,gn6, ncol=3)
gridExtra::grid.arrange(gn7,gn8,gn9, ncol =3)
gridExtra::grid.arrange(gn10,gn11, ncol =3 )

#____________________________________________________________________

#val = data$Distance.from.Residence.to.Work[data$Distance.from.Residence.to.Work %in% boxplot.stats(data$Distance.from.Residence.to.Work)$out]
#val

for (i in n_data) {
  print(i)
  val = data[,i][data[,i] %in% boxplot.stats(data[,i])$out]
  print(length(val))
  print(val)
}

#Make each outlier as NA
for (i in n_data) {
  val = data[,i][data[,i] %in% boxplot.stats(data[,i])$out]
  data[,i][data[,i] %in% val] = NA
}

#Check number of missing values
sum(is.na(data))

#Get number of missing values in each variable
for (i in n_data) {
  print(i)
  print(sum(is.na(data[,i])))
}

#Impute the values using KNN method
data = knnImputation(data, k=3)

#Check again for missing value if present in case
sum(is.na(data))

#Now check for outlier in case if generated after KNN imputation
for (i in 1:length(n_data)) {
      print(i)
      assign(paste0("g",i), ggplot(aes_string(y = (n_data[i]), x= "Absenteeism.time.in.hours"), data=subset(data))+
      stat_boxplot(geom = "errorbar", width= 0.5) +
      geom_boxplot(outlier.color = "red" , fill = "grey", outlier.shape = 18, outlier.size = 1, notch = FALSE) +
      theme(legend.position = "bottom")+
      labs(y = n_data[i], x="Absenteeism.time.in.hours") +
      ggtitle(paste("Box Plot of Employee data", n_data[i])))
         
}

gridExtra::grid.arrange(g1,g2, ncol = 2) #no outlier seen 
gridExtra::grid.arrange(g3,g4,g5, ncol = 3) #no outlier seen
gridExtra::grid.arrange(g6,g7,g8, ncol=3) #no outlier seen
gridExtra::grid.arrange(g9,g10,g11, ncol=3) #no outlier seen

#Confirm using boxplot.stat method to see whether outlier exists
for (i in n_data) {
  print(i)
  val = data[,i][data[,i] %in% boxplot.stats(data[,i])$out]
  print(length(val))
  print(val)
}

#No outliers found

#Correlation plot
library(corrgram)
corrgram(na.omit(data))
dim(data)
corrgram(data[,n_data],order = F, upper.panel = panel.pie, text.panel = panel.txt, main = "correlation plot" )

#Here we can see in correlation plot that BMI and weight are highly positively correlated
# I am removing the BMI variable because weight is a basic variable.



# Select the relevant numerical features

num_vars = c("Transportation.expense", "Distance.from.Residence.to.work", "Service.time", "Age", "Work.load.Average.day","Hit.target", "Son", "Pet", "Weight", "Height")
n_data
names(data)
#install.packages("lsr")

library("lsr")

anova_test = aov(Absenteeism.time.in.hours ~ ID + Day.of.the.week + Education + Social.smoker + Social.drinker + Reason.for.absence + Seasons + Month.of.absence + Disciplinary.failure, data = data)
summary(anova_test)

#Dimension reduction
data_selected = subset(data,select = -c(ID,Body.mass.index, Education, Social.smoker, Social.drinker, Seasons, Disciplinary.failure))

#Make a backup of features selected

feature_selected = data_selected
write.csv(data_selected,"Selected_Features.csv", row.names = F)

colnames(data_selected)

#draw histogram on random variables to check if the distributions are normal

hist(data_selected$Hit.target)
hist(data_selected$Work.load.Average.day.)
hist(data_selected$Son)
hist(data_selected$Weight)
hist(data_selected$Transportation.expense)
hist(data_selected$Distance.from.Residence.to.Work)
hist(data_selected$Service.time)
hist(data_selected$Age)
hist(data_selected$Pet)


#THE variables are not seems as normally distributed thus we have to perform Normalisation instead of standardisation
#Now select numerical variables from data_selected Object to perform normalization

print(sapply(data_selected ,is.numeric))
num_names = c("Transportation.expense", "Distance.from.Residence.to.Work","Service.time","Age","Work.load.Average.day.","Hit.target","Son","Pet","Weight","Height")

#Feature Scaling
#num_names object contains all the numerical features of selected features

for (i in num_names) {
     print(i)
     data_selected[,i] = ((data_selected[,i] - min(data_selected[,i])) /
                          (max(data_selected[,i]) - min(data_selected[,i])))
   
}
colnames(data_selected)

#Till here we are having our data being normalised and stored in data_selected object
#Write this cleaned data in csv form 
write.csv(data_selected,"clean_data_selected.csv", row.names = F)

#________________________________________________________________________________
library(DataCombine)
rmExcept(c("data_selected","feature_selected"))
colnames(data_selected)
View(data_selected)

#___________ MODEL DEVELOPMENT _______________________

## Divide the data into train and test data using simple random sampling

library("rpart")
train_index = sample(1:nrow(data_selected), 0.8* nrow(data_selected))

train = data_selected[train_index,]
test = data_selected[-train_index,]


#Building the Decision Tree Regression Model


#reg_model = rpart(Absenteeism.time.in.hours ~. , data = train, method = "anova")

#########  predicting for the test case
#predicted_data = predict(reg_model , test[,-13])

##Calculate RMSE to analyze performance of Decision tree regression model
#RMSE is used here as the data is of time series 

#library("DMwR")
#regr.eval(test[,13], predicted_data, stats = 'rmse')

##__ Thus in Decision tree regression model the error is 9.10 which tells that our model is 90.9% accurate

##________________RANDOM FOREST MODEL __________________________________##

library("randomForest")
RF_model = randomForest(Absenteeism.time.in.hours~. , train,  ntree=100)

#Extract the rules generated as a result of random Forest model
library("inTrees")
rules_list = RF2List(RF_model)

#Extract rules from rules_list
rules = extractRules(rules_list, train[,-14])
rules[1:2,]

#Convert the rules in readable format
read_rules = presentRules(rules,colnames(train))
read_rules[1:2,]

#Determining the rule metric
rule_metric = getRuleMetric(rules, train[,-14], train$Absenteeism.time.in.hours)
rule_metric[1:2,]

#Prediction of the target variable data using the random Forest model
RF_prediction = predict(RF_model,test[,-14])
regr.eval(test[,13], RF_prediction, stats = 'rmse')

#Thus the error rate in Random Forest Model is 8.78% and the accuracy of the model is 100-8.78 = 91.22%

###________________ LINEAR REGRESSION _______________________________

#library("usdm")
#LR_data_select = subset(data_selected ,select = -c(Reason.for.absence,Day.of.the.week))
#colnames(LR_data_select)
#vif(LR_data_select[,-11])
#vifcor(LR_data_select[,-11], th=0.9)

####Execute the linear regression model over the data
#lr_model = lm(Absenteeism.time.in.hours~. , data = train)

#summary(lr_model)
#colnames(test)

#Predict the data 
#LR_predict_data = predict(lr_model, test[,1:12])

#Calculate MAPE
#MAPE(test[,13], LR_predict_data)
#library("Metrics")
#rmse(test[,13],LR_predict_data)

##________ Till here we have implemented Decision Tree, Random Forest and Linear Regression. Among all of these,
##________ Random Forest is having highest accuracy.


#________ Now we have to predict loss for each month ____________

#__ To calculate loss month wise we need to include month of absence variable again in our data set 
# LOSS = Work.load.average.per.day * Absenteeism.time.in.hours

data = feature_selected

colnames(data)
data$loss = data$Work.load.Average.day. * data$Absenteeism.time.in.hours

i=1

#______________________________________________________________________________

# NOW calculate Month wise loss encountered due to absenteeism of employees 

#Calculate loss in january
loss_jan = as.data.frame(data$loss[data$Month.of.absence %in% 1])
names(loss_jan)[1] = "Loss"
sum(loss_jan[1])
write.csv(loss_jan,"jan_loss.csv", row.names = F)

#Calculate loss in febreuary
loss_feb = as.data.frame(data$loss[data$Month.of.absence %in% 2])
names(loss_feb)[1] = "Loss"
sum(loss_feb[1])
write.csv(loss_feb,"feb_loss.csv", row.names = F)

#Calculate loss in march

loss_march = as.data.frame(data$loss[data$Month.of.absence %in% 3])
names(loss_march)[1] = "Loss"
sum(loss_march[1])
write.csv(loss_march,"march_loss.csv", row.names = F)

#Calculate loss in april

loss_apr = as.data.frame(data$loss[data$Month.of.absence %in% 4])
names(loss_apr)[1] = "Loss"
sum(loss_apr[1])
write.csv(loss_apr,"apr_loss.csv", row.names = F)


#calculate loss in may

loss_may = as.data.frame(data$loss[data$Month.of.absence %in% 5])
names(loss_may)[1] = "Loss"
sum(loss_may[1])
write.csv(loss_may,"may_loss.csv", row.names = F)


#calculate in june

loss_jun = as.data.frame(data$loss[data$Month.of.absence %in% 6])
names(loss_jun)[1] = "Loss"
sum(loss_jun[1])
write.csv(loss_jun,"jun_loss.csv", row.names = F)

#Calculate loss in july

loss_jul = as.data.frame(data$loss[data$Month.of.absence %in% 7])
names(loss_jul)[1] = "Loss"
sum(loss_jul[1])
write.csv(loss_jul,"jul_loss.csv", row.names = F)

#calculate loss in august

loss_aug = as.data.frame(data$loss[data$Month.of.absence %in% 8])
names(loss_aug)[1] = "Loss"
sum(loss_aug[1])
write.csv(loss_aug,"aug_loss.csv", row.names = F)

#Calculate loss in september

loss_sep = as.data.frame(data$loss[data$Month.of.absence %in% 9])
names(loss_sep)[1] = "Loss"
sum(loss_sep[1])
write.csv(loss_sep,"sep_loss.csv", row.names = F)

#calculate loss in october

loss_oct = as.data.frame(data$loss[data$Month.of.absence %in% 10])
names(loss_oct)[1] = "Loss"
sum(loss_oct[1])
write.csv(loss_oct,"oct_loss.csv", row.names = F)

#calculate loss in november

loss_nov = as.data.frame(data$loss[data$Month.of.absence %in% 2])
names(loss_nov)[1] = "Loss"
sum(loss_nov[1])
write.csv(loss_nov,"nov_loss.csv", row.names = F)

#calculate loss in december
loss_dec = as.data.frame(data$loss[data$Month.of.absence %in% 2])
names(loss_dec)[1] = "Loss"
sum(loss_dec[1])
write.csv(loss_dec,"dec_loss.csv", row.names = F)

