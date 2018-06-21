# install.packages("readxl")
# install.packages("ggplot2")
# install.packages("labeling")
# install.packages("reshape2")
# install.packages("data.table")
# install.packages("psych")
# install.packages("pastecs")
# install.packages("Hmisc")
# install.packages("BSDA")
# install.packages("magrittr")
# install.packages("stringr")
#1 reading excel file----

library(readxl)
library(ggplot2)
library(labeling)
library(reshape2)
library(data.table)
library(psych)
library(pastecs)
library(Hmisc)
library(BSDA)
library(magrittr)
library(stringr)
sink('Chapter_1.txt')
#path = "C:\Users\ishan\Desktop\Lab1\Performance Lawn Equipment Database.xlsx"
#path = "C:/Users/ishan/Desktop/Lab1/Performance Lawn Equipment Database.xlsx"
path = "Desktop/Harrisburg/ANLY 500 - Prin of Analy/Performance Lawn Equipment Database.xlsx"
excelFile = read_excel(path)
#View(excelFile)

sheets = readxl::excel_sheets(path)
size = length(sheets)
excelData = list()
for(i in seq_along(sheets)){
  sheet = sheets[i]
  excel = read_excel(path,sheet,skip =2)
  excelData[[i]] = data.frame(excel)
  print(sheet)
  str(excelData[i])
}
sink()

#3 chapter 3----data-
#3.1 Dealer satisfaction charts from tutorialD----
countries = c("North America", "South America", "China","Europe", "Pacific Rim")
k =0
for(j in seq_along(excelData))
{
  
  for(i in seq_along(countries))
  {
    k= 0;
    frame = excelData[[j]]
    dealerSatisfaction = t(frame[which(frame$Region == countries[i]),3:8])
    dealerSatisfaction = t(excelData[[j]][1:5,3:8])
    print(dealerSatisfaction)
    print (i)
    colnames(dealerSatisfaction) = c("2010","2011","2012","2013","2014")
    data.m2 = melt(dealerSatisfaction, id.vars = Var1)
    print(data.m2)
    colnames(data.m2) = c("Levels","Years","Counts")
    
    ggplot(data.m2, aes(x=countries[i], y=Counts)) +
           geom_bar(aes(fill=Levels), position="dodge", stat="identity") +
           ggtitle(sheets[j])

    fname <- sprintf('%s_',deparse(sheets[j]))
    f1 = paste(paste(fname,k),".jpg")
    str_replace_all(string=f1, pattern=" ", repl="")
    ggsave(f1,last_plot())
    
    k =k + 1
     print(ggplot(data.m2, aes(x=Years, y=Counts, fill=Levels)) +
           geom_bar(stat="identity") +
           ggtitle(sheets[[j]]))
  
     ggplot(data.m2, aes(x=Years, y=Counts)) + geom_bar(aes(fill=Levels), position="dodge", stat="identity")
     fname <- sprintf('%s_',deparse(sheets[j]))
     f1 = paste(paste(fname,k),".jpg")
     str_replace_all(string=f1, pattern=" ", repl="")
     ggsave(f1,last_plot())
     
     fname <- sprintf('%s_',deparse(sheets[j]))
     f1 = paste(paste(fname,k),".jpg")
     str_replace_all(string=f1, pattern=" ", repl="")
     ggsave(f1,last_plot())
    ggsave("Years vs Count-Bar.jpg",last_plot())

    k= k+1
     plot = ggplot(data.m2, aes(x=Years, y=Counts, fill=Levels)) +
           geom_line(aes(color=Levels),stat="identity") +
           geom_point(stat="identity") +
           ggtitle(sheets[[j]])
   
     fname <- sprintf('%s_',deparse(sheets[j]))
     f1 = paste(paste(fname,k),".jpg")
     str_replace_all(string=f1, pattern=" ", repl="")
     ggsave(f1,last_plot())

  }
  if(j == 2)
    break
}

sink('Chapter_3.txt')


# Step 2
Complaints = excelData[[4]]

plot(Complaints$World, ylim=range(c(0,400)), type="l", xlab="Month", ylab="Number of Complaints")
par(new=TRUE)

plot(Complaints$NA., ylim=range(c(0,400)), type="l", col="red", axes = FALSE, xlab = "", ylab = "")
par(new=TRUE)

plot(Complaints$SA, ylim=range(c(0,400)), type="l", col="green", axes = FALSE, xlab = "", ylab = "")
par(new=TRUE)
 
plot(Complaints$Eur, ylim=range(c(0,400)), type="l", col="blue", axes = FALSE, xlab = "", ylab = "")
par(new=TRUE)

plot(Complaints$Pac, ylim=range(c(0,400)), type="l", col="magenta", axes = FALSE, xlab = "", ylab = "")
par(new=TRUE)

plot(Complaints$China, ylim=range(c(0,400)), type="l", col="deeppink4", axes = FALSE, xlab = "", ylab = "")




#3.2 Summary command-----
shippingCost = excelData[[21]]
summary(shippingCost$Mowers)
summary(shippingCost$Tractors)


#3.3 2014 Customer survey----
CustomerSurvey2014 =  read_excel(path,sheets[[3]],skip =1)
barplot(table(CustomerSurvey2014$Region),main = "Customer Surver - Number of response",xlab = "Region", ylab = "Number of responses")
hist(CustomerSurvey2014$Quality, main="Quality - Number of Responses", xlab="Level of Quality")
summary(CustomerSurvey2014)
sink()

#3.4 Most important business?----

#4.a Mean and standard deviation for end user & dealer satisfaction-----
sink('Chapter_4.txt')
for(j in seq_along(excelData))
{
  dealerSatisfaction = excelData[[j]]
  mean =0;
  dev =0;
  for(i in 1:25)
  {
    mean[i]= (
      1*(dealerSatisfaction[i,4]) + 2*(dealerSatisfaction[i,5]) +
        3*(dealerSatisfaction[i,6]) + 4*(dealerSatisfaction[i,7]) +
        5*(dealerSatisfaction[i,8]))/sum(dealerSatisfaction[i,3:8])
    
    
    dev[i] = sqrt(((dealerSatisfaction[i,3] * (0 - mean[i])^2) + 
                     (dealerSatisfaction[i,4] * (1 - mean[i])^2) +
                     (dealerSatisfaction[i,5] * (2 - mean[i])^2) +
                     (dealerSatisfaction[i,6] * (3 - mean[i])^2) +
                     (dealerSatisfaction[i,7] * (4 - mean[i])^2) +
                     (dealerSatisfaction[i,8] * (5 - mean[i])^2))/
                    (sum(dealerSatisfaction[i,3:8])-1))
    
  }
  n = matrix(mean, 5, byrow = FALSE)
  print("MEAN VALUES: \n")
  print(n)
  
  d = matrix(dev, 5, byrow = FALSE)
  print("STANDARD DEVIATION VALUES: \n")
  
  print(d)
  print("\n")
  if(j ==2)
    break
}


#4.b Descriptive stats using psych package-----
countries = c("NA", "SA", "China","Eur", "Pac")
print("#4.b Descriptive stats using psych package-----")
for(i in seq_along(countries))
{
  cust2014= subset(CustomerSurvey2014, subset = CustomerSurvey2014$Region == countries[i])
  print(countries[i])
  print(describe(cust2014))
  print(stat.desc(cust2014))
}
sink()
#4.c Response Time ----
reTime = excelData[[14]]
str(reTime)
meanResponseTime = c(mean(reTime$Q1.2013), mean(reTime$Q2.2013), mean(reTime$Q3.2013), mean(reTime$Q4.2013)
                     , mean(reTime$Q1.2014), mean(reTime$Q2.2014), mean(reTime$Q3.2014), mean(reTime$Q4.2014))

plot(meanResponseTime, type = "l", lwd=2, col="green", ylab = "Mean", xlab = "Quarter", 
     main="Mean Response Time", ylim = c(min(meanResponseTime),max(meanResponseTime)), xaxt="n")

axis(side = 1, at = c(1:8), labels = colnames(reTime), pch=0.5)


#4.d Defects Over Delivery and plot against time-----

defects = excelData[[12]]
str(defects)
meanDefects = c(mean(defects$X__1), mean(defects$X__2), mean(defects$X__3)
                ,mean(defects$X__4), mean(defects$X__5))

plot(meanDefects, type = "l", lwd=2, col="green", ylab = "Mean", xlab = "Quarter", 
     main="Mean Defects", ylim = c(min(meanDefects),max(meanDefects)), xaxt="n")

axis(side = 1, at = c(1:5), labels = c("2010","2011","2012","2013","2014"), pch=0.5)

#4.e Mower and Industry (index 5,7) comparison ERROR: DATA IS WRONG: NUMBER OF COLUMN IS LESS
# Tractor are index 6 and 8
sink('Chapter_4.txt',append = TRUE)

i =4
for(j in 1:2)
{
  mowerData = excelData[[i+j]]
  industryData = excelData[[i+j+2]]
  
  #str(mowerData)
  
  #str(industryData)
  
  totalData = cbind(mowerData[,],industryData[,-1])
  colnames(totalData) = c("Month" , "NorthA", "SA", "Eur", "Pac", "China", "World",
                          "IndusNA", "IndSA", "IndEur", "IndPac", "IndWorld")
  
  stat.desc(totalData)
  
  rCor = rcorr(as.matrix(totalData[2:12]))
  
  r_matrix = as.matrix(rCor$r)
  
  corrplot::corrplot(r_matrix,title = sheets[[i+j]],win.asp = 0.8)
  
}

sink()

#5.1 ----- Bernoulli Distribution

sink('Chapter_5.txt')
#5.2 -----
mowerTest = excelData[[19]]

countFail = length(which(mowerTest == "Fail"))
countPass = length(which(mowerTest == "Pass"))
probF = countFail /(countPass + countFail)
print(probF)


#5.3 -------

probXF = dbinom(0:20,100,probF)
print(probXF)
plot(probXF, main = "Binomial distribution")

#5.4 -----

bladeWeight = excelData[[18]]
colnames(bladeWeight) = c("sample","weight")
avgW = sum(bladeWeight$weight)/length(bladeWeight$weight)
print(avgW)

var = sd(bladeWeight$weight)
print(var)

#5.5 more than 5.20---

probMoreThan5 = pnorm(5.20,mean = avgW, sd = var)
print(1-probMoreThan5)

#5.6 Less than 4.80

problessThan4 = pnorm(4.80,mean = avgW, sd = var)
print(problessThan4)

#5.7 blades excedding ----
countBlade1 = length(which(bladeWeight$weight > 5.20))
print(countBlade1)
countBlade2 = length(which(bladeWeight$weight <= 4.80))
print(countBlade2)

#5.8 Is the process stable

plot(bladeWeight$sample, bladeWeight$weight, main ="Changes in blade weight over time (Scatter-plot)")

#5.9 Count number of outliners 172

#5.10 normal distribution

histPlot = hist(bladeWeight$weight,ylim = c(0 ,6), prob = TRUE)
curve(dnorm(x, mean = avgW , sd = var), col = "darkblue",lwd = 2, add=TRUE)

sink()
#6.1 ------

sink('Chapter_6.txt')



CustomerSurvey2014 = excelData[[3]]
countRegion = table(CustomerSurvey2014$Region, CustomerSurvey2014$Quality)
result = sum(countRegion[3,4], countRegion[3,5])
print(result)

#6.2-------
reTime = excelData[[14]]

#From 4.2c
meanResponseTime = c(mean(reTime$Q1.2013), mean(reTime$Q2.2013), mean(reTime$Q3.2013), mean(reTime$Q4.2013)
                     , mean(reTime$Q1.2014), mean(reTime$Q2.2014), mean(reTime$Q3.2014), mean(reTime$Q4.2014))
stdResponseTime =  c(sd(reTime$Q1.2013), sd(reTime$Q2.2013), sd(reTime$Q3.2013), sd(reTime$Q4.2013)
                     , sd(reTime$Q1.2014), sd(reTime$Q2.2014), sd(reTime$Q3.2014), sd(reTime$Q4.2014))

for(i in seq_along(reTime)){
  print(z.test(reTime[,i], sigma.x = stdResponseTime[i]))
}

#6.3 ------
transmissionCost = excelData[[17]]

meanC = c(mean(transmissionCost$Current), mean(transmissionCost$Process.A), mean(transmissionCost$Process.B))
sdC = c(sd(transmissionCost$Current), sd(transmissionCost$Process.A), sd(transmissionCost$Process.B))

for(i in seq_along(transmissionCost)){
  print(z.test(transmissionCost[,i], sigma.x = sdC[i]))
  if(i ==3)
    break
  
}

#6.4-----
mowerTest = read_excel(path, sheet = "Mower Test", skip = 3)

A = data.frame(lapply(mowerTest, function(x) as.numeric(as.factor(x))))

#str(A)
A[,5] = 2
A[,10] = 2
A[,27] = 2
A[,30] = 2

#str(A)
B =data.matrix(A[,-1])
B = B - 1

length(which(B == 0))
mean(B)
1-mean(B)

t.test(B, mu=mean(B))

lowerFail = 1 - 0.9772398
highFail = 1 - 0.9867602

#6.5----
hist(bladeWeight$weight)
em = sd(bladeWeight$weight)/sqrt(length(bladeWeight$weight))
A = lm( bladeWeight$weight ~ ., bladeWeight)
plot(A)



#Chapter 7------
#7.1------
Region = CustomerSurvey2014$Region
Quality = CustomerSurvey2014$Quality
Ease =  CustomerSurvey2014$Ease.of.Use
Price = CustomerSurvey2014$Price
Service = CustomerSurvey2014$Service

Data = data.frame(Y=c(Quality, Ease, Price, Service), 
                   Para=factor(rep(c("Quality", "Ease.of.Use", "Price", "Service"), 
                  times=c(length(Quality), length(Ease), length(Price), length(Service)))))

m1 = aov(Y~Para, data = Data)
anova(m1)



#7.2------
OnTimeDelivery = excelData[[11]]
p0 <- sum(OnTimeDelivery$Number.On.Time[1:12])/sum(OnTimeDelivery$Number.of.deliveries[1:12])

pbar <- sum(OnTimeDelivery$Number.On.Time[49:60])/sum(OnTimeDelivery$Number.of.deliveries[49:60])
n =12
zScore = (pbar-p0)/sqrt(p0*(1-p0)/n)

pVal = pnorm(zScore, lower.tail = FALSE)

print(pVal)


#7.3------

DefectsAfterDelivery = excelData[[12]]
colnames(DefectsAfterDelivery) = c("Month","2010","2011","2012","2013","2014")
DefectsAfterDelivery = DefectsAfterDelivery[-1,]
sum(DefectsAfterDelivery$`2014`)/12
defects.by.year <- c(826.33, 837.42, 785.92, 669.08, 496.25)
plot(defects.by.year, type="b", lwd=2,main ="Defects After Delivery for 2014")

t.test(DefectsAfterDelivery$`2010`, DefectsAfterDelivery$`2014`, var.equal = FALSE, paired = FALSE)

#7.4-----
t.test(transmissionCost$Current, transmissionCost$Process.A, var.equal = FALSE, paired = FALSE)

t.test(transmissionCost$Current, transmissionCost$Process.B, var.equal = FALSE, paired = FALSE)


#7.5 -----

employeeRetension = excelData[[20]]

gender = employeeRetension$Gender
yearsPLE = employeeRetension$YearsPLE

male.female = data.frame(gender=gender, yearsPLE=yearsPLE)

male.female[order(male.female$gender),]

t.test(male.female$yearsPLE ~ male.female$gender, var.equal = FALSE, paired = FALSE)



