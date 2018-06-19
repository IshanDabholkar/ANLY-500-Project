#Project

#----Read the file path
library(readxl)
library(psych)
library(ggplot2)

path = "C:/Users/ishan/Desktop/HU/ANLY-500-Project-master/ANLY-500-Project-master/CustomerChurn.xlsx"
sheets = readxl::excel_sheets(path)
size = length(sheets)
excelData = list()

for(i in seq_along(sheets)){
  sheet = sheets[i]
  excel = read_excel(path,sheet,skip =0)
  excelData[[i]] = data.frame(excel)
  print(sheet)
  str(excelData[i])
}

customerChurn = excelData[[1]]

#----- Basic descriptive analysis
summary(customerChurn)
describe(customerChurn)

#------ Begin some data filtering and analysis
customerChurn_No = subset(customerChurn , Churn == "No")
customerChurn_Yes = subset(customerChurn , Churn == "Yes")

j = 1
#------ Total number
plot1 = ggplot(customerChurn, aes(x=customerChurn$PaymentMethod, y=customerChurn$MonthlyCharges , fill = PaymentMethod)) +
    geom_bar(stat = "identity") +
    ggtitle(sheets[j])

#------- When customer churn is "No"
plot2 = ggplot(customerChurn_No, aes(x=customerChurn_No$PaymentMethod, y=customerChurn_No$MonthlyCharges , fill = PaymentMethod)) +
  geom_bar(stat = "identity") +
  ggtitle(sheets[j])



#----- Conclusion one - Maybe people who churn use electronic checks for monthly payments a lot
plot3 = ggplot(customerChurn_Yes, aes(x=customerChurn_Yes$PaymentMethod, y=customerChurn_Yes$MonthlyCharges , fill = PaymentMethod)) +
    geom_bar(stat = "identity") +
  ggtitle(sheets[j])



#------- Senior Citizen and Churn or not?
#--------Conclusion two- Maybe poeple who churn are more likely to be not senior citizen
plot4 =ggplot(customerChurn, aes(x = customerChurn$Churn, y = customerChurn$SeniorCitizen , fill = Churn)) + 
    geom_bar(stat = "identity") 

#------- Plot in a PIE CHART ????-- Need to plot Stacked bar chart   pieChart = plotChurn_Senior + coor
  

#------ Conculusion three - Male are more likely to churn than female 
plot5 = ggplot(customerChurn_Yes, aes(x = customerChurn_Yes$Churn, y = customerChurn_Yes$gender , fill = gender)) + 
  geom_bar(stat = "identity") 


#------- What services do people who churn use mostly
#---- PhoneServe, InternetService, OnlinSecurity, OnlineBackup, DeviceProtection


numeric_Phone = as.numeric(as.factor(customerChurn$PhoneService))
numeric_OnlineSec = as.numeric(as.factor(customerChurn$OnlineSecurity))
numeric_InternetServ = as.numeric(as.factor(customerChurn$InternetService))

ggplot(customerChurn, aes(x = numeric_Phone, fill = PhoneService)) + 
  geom_bar(position = "dodge") 


#print(plot6)

print(plot1)
print(plot2)
print(plot3)

print(plot4)

print(plot5)


#----- Define a function to plot the histogram and the normal curve
plotHistogram = function(x, title_, ...)
{
  
  x = x[!is.na(x)]
  
  distNorm_MonthlyCharges = dnorm(x, mean = mean(x) , sd = sd(x))
  
  h = hist(x, xlab = "Monthly charges", ylab = "Frequency", main = title_)
  
  par(new = TRUE)
  
  xfit = seq(min(x),max(x),length= length(x) )
  
  yfit = dnorm(xfit,mean=mean(x),sd=sd(x)) 
  
  yfit = yfit* diff(h$mids[1:2]) * length(x)
  
  lines(xfit, yfit, col="blue", lwd=2)
  
}

#------ Distribution of monthly charges 
plotHistogram(customerChurn$MonthlyCharges, "Distribution of Monthly Charge")
 
#------ Distribution of Total charges 
plotHistogram(customerChurn$TotalCharges, "Distribution of Total Charge")

#------ Distribution of Total charges 
plotHistogram(customerChurn_Yes$tenure, "Distribution of Tenure for poeple who Churn")

## Which type of customer are more likely to churn
# Contract type - Month to month, One year, two year
# Conclusion 4 - MONTH TO MONTH are more likely to churn

x = customerChurn$Contract
x = x [!is.na(x)]
ggplot(customerChurn , aes(x = x, y = customerChurn$Churn, fill = Churn)) +geom_bar( stat = "identity")

## Conclusion 5- Which type of people are likely to churn?
## PhoneService, InternetService, OnlineBackup etc..
## Infer from Pie charts?
## May be we need to remove the "No internet service" for better percentage analysis
pie(table(customerChurn_No$InternetService))
pie(table(customerChurn_Yes$InternetService))

pie(table(customerChurn_No$PhoneService))
pie(table(customerChurn_Yes$PhoneService))

pie(table(customerChurn_No$OnlineBackup))
pie(table(customerChurn_Yes$OnlineBackup))

pie(table(customerChurn_No$OnlineSecurity))
pie(table(customerChurn_Yes$OnlineSecurity))




