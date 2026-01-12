#загружаем библиотеки 
library("miceadds")
library("clubSandwich")
library("sandwich") 
library("ggcorrplot")
library("dplyr")
library("ggplot2")
library("GGally")
library("vars")
library("urca")
library("TSA")
library("forecast")
library("lmtest")
library("tseries")
library("urca")
library("TSA")
library("openxlsx")
library("plm")
library("car")
library("corrplot")
library("Hmisc")

#строим модели для нахождения AEM и сравниваем их 
d <- read.xlsx ('a.xlsx',sheet = "oil")
d <- na.omit(d)
m.re1 <- plm(data = d, d$`TA/At-1`~d$`1/At-1`+d$`dR-dR/At-1`+d$`PPE/At-1`+d$ROA, model = "random")
summary(m.re1)
bptest(m.re1)
coeftest(m.re1, vcov. =vcovHC, type = 'HC0') 

#оцениваем NDA для горно-металлургической отрасли
k <- read.xlsx ('a.xlsx',sheet = "metal")
k<-na.omit(k)
m.pooled2 <- plm(data = k, k$`TA/At-1`~k$`1/At-1`+k$`dR-dR/At-1`+ k$`PPE/At-1` +k$ROA, model = "pooling")
summary(m.pooled2)
bptest(m.pooled2)
coeftest(m.pooled2, vcov. =vcovHC, type = 'HC0')

#оцениваем NDA для энергетической отрасли
r <- read.xlsx ('a.xlsx',sheet = "energy")
r<-na.omit(r)
m.pooled3 <- plm(data = r, r$`TA/At-1`~r$`1/At-1`+r$`dR-dR/At-1`+r$`PPE/At-1`+r$ROA, model = "pooling")
summary(m.pooled3)
bptest(m.pooled3)
coeftest(m.pooled3, vcov. =vcovHC, type = 'HC0')

#коэффициенты лучших моделей
summary(m.re1)
summary(m.pooled2)
summary(m.pooled3)

#оцениваем CFO нефть
e <- read.xlsx ('a.xlsx',sheet = "remoilcfo")
e <- na.omit(e)
m.pooled4cfonorm <- plm(data = e, e$`CFO/A(t-1)`~e$`1/At-1`+e$`Sales/A(t-1)`+e$`deltaSales/At-1`, model = "pooling")
summary(m.pooled4cfonorm)
bptest(m.pooled4cfonorm)

#оцениваем CFO металл
mm <- read.xlsx ('a.xlsx',sheet = "remmetalcfo")
mm <- na.omit(mm)
m.re5cfonorm <- plm(data = mm, mm$`CFO/A(t-1)`~mm$`1/At-1`+mm$`Sales/A(t-1)`+mm$`deltaSales/At-1`, model = "random")
summary(m.re5cfonorm)
bptest(m.re5cfonorm)

#оцениваем CFO энергия
m <- read.xlsx ('a.xlsx',sheet = "remenergycfo")
m <- na.omit(m)
m.re6cfonorm <- plm(data = m, m$`CFO/A(t-1)`~m$`1/At-1`+m$`Sales/A(t-1)`+m$`deltaSales/At-1`, model = "random")
summary(m.re6cfonorm)
bptest(m.re6cfonorm)

#оцениваем PROD нефть
e <- read.xlsx ('a.xlsx',sheet = "remoilprod")
e <- na.omit(e)
m.pooled4prodnorm <- plm(data = e, e$`PROD/At-1`~e$`1/At-1`+e$`Sales/A(t-1)`+e$`deltaSales/At-1`+e$`deltadelta/At-1`, model = "pooling")
summary(m.pooled4prodnorm)
bptest(m.pooled4prodnorm)

#оцениваем PROD металл
mm <- read.xlsx ('a.xlsx',sheet = "remmetalprod")
mm <- na.omit(mm)
m.pooled5prodnorm <- plm(data = mm, mm$`PROD/At-1`~mm$`1/At-1`+mm$`Sales/A(t-1)`+mm$`deltaSales/At-1`+mm$`deltadelta/At-1`+mm$`deltadelta/At-1`, model = "pooling")
summary(m.pooled5prodnorm)
bptest(m.pooled5prodnorm)
coeftest(m.pooled5prodnorm, vcov. =vcovHC, type = 'HC0')

#оцениваем PROD энергия
m <- read.xlsx ('a.xlsx',sheet = "remenergyprod")
m <- na.omit(m)
m.re6prodnorm <- plm(data = m, m$`PROD/At-1`~m$`1/At-1`+m$`Sales/A(t-1)`+m$`deltaSales/At-1`+m$`deltadelta/At-1`, model = "random")
summary(m.re6prodnorm)
bptest(m.re6prodnorm)

#корреляционные матрицы
e <- read.xlsx ('a.xlsx',sheet = "corpnnp")
e <- na.omit(e)
ggcorr(e, method = c("pairwise", "pearson"),
       nbreaks = NULL, digits = 2, low = "#3B9AB2",
       mid = "#EEEEEE", high = "#F21A00",
       geom = "tile", label = FALSE,
       label_alpha = FALSE)

o <- read.xlsx ('a.xlsx',sheet = "corpublic")
o <- na.omit(o)
ggcorr(o, method = c("pairwise", "pearson"),
       nbreaks = NULL, digits = 2, low = "#3B9AB2",
       mid = "#EEEEEE", high = "#F21A00",
       geom = "tile", label = FALSE,
       label_alpha = FALSE)
ggcorr(o, palette = "RdBu", label = TRUE)

#описательные статистики переменных для выборки публичных и непубличных компаний 
q <- read.xlsx ('a.xlsx',sheet = "pnnp")
q <- na.omit(q)
summary(q)

#описательные статистики переменных для выборки публичных компаний
p <- read.xlsx ('a.xlsx',sheet = "public")
p <- na.omit(p)
summary(p)

#сравниваем распределение переменных с нормальным
#публичные и непубличные компаниий

q <- read.xlsx ('a.xlsx',sheet = "pnnp")
q <- na.omit(q)

hist(q$AEEM,
     xlab = "AEM", 
     ylab = "Частота", 
     main = "AEM",
     col = "plum",
     freq = F 
)

mean_beaver <- mean(q$AEM)
sd_beaver <- sd(q$AEM)

curve(dnorm(x,
            mean = mean_beaver, 
            sd = sd_beaver),
      add = T 
)

hist(q$REM,
     xlab = "REM", 
     ylab = "Частота", 
     main = "REM",
     col = "plum", 
     freq = F 
)
mean_beaver1 <- mean(q$REM)
sd_beaver1 <- sd(q$REM)

curve(dnorm(x,
            mean = mean_beaver1, 
            sd = sd_beaver1),
      add = T 
)

hist(q$GROWTH,
     xlab = "GROWTH", 
     ylab = "Частота", 
     main = "GROWTH",
     col = "plum", 
     freq = F 
)
mean_beaver2 <- mean(q$GROWTH)
sd_beaver2 <- sd(q$GROWTH)

curve(dnorm(x,
            mean = mean_beaver2, 
            sd = sd_beaver2),
      add = T 
)

hist(q$ROAadj,
     xlab = "ROAadj", 
     ylab = "Частота", 
     main = "ROAadj",
     col = "plum", 
     freq = F 
)
mean_beaver3 <- mean(q$ROAadj)
sd_beaver3 <- sd(q$ROAadj)

curve(dnorm(x,
            mean = mean_beaver3, 
            sd = sd_beaver3),
      add = T 
)

hist(q$SIZE,
     xlab = "SIZE", 
     ylab = "Частота", 
     main = "SIZE",
     col = "plum", 
     freq = F 
)
mean_beaver4 <- mean(q$SIZE)
sd_beaver4 <- sd(q$SIZE)

curve(dnorm(x,
            mean = mean_beaver4, 
            sd = sd_beaver4),
      add = T 
)

hist(q$E_SCORE,
     xlab = "E_SCORE", 
     ylab = "Частота", 
     main = "E_SCORE",
     col = "plum", 
     freq = F 
)
mean_beaver5 <- mean(q$E_SCORE)
sd_beaver5 <- sd(q$E_SCORE)

curve(dnorm(x,
            mean = mean_beaver5, 
            sd = sd_beaver5),
      add = T 
)

#избавляемся от выбросов
q <- read.xlsx ('a.xlsx',sheet = "pnnp")
q <- na.omit(q)

outliers <- boxplot(q$REM, plot=FALSE)$out
ol <- q[-which(q$REM %in% outliers),]
q<-ol
outliers <- boxplot(q$GROWTH, plot=FALSE)$out
ol <- q[-which(q$GROWTH %in% outliers),]
q<-ol

#тесты на нормальность распределения
shapiro.test(q$AEM)
shapiro.test(q$REM)
shapiro.test(q$ROAadj)
shapiro.test(q$SIZE)
shapiro.test(q$GROWTH)
shapiro.test(q$E_SCORE)

#публичные компании
p <- read.xlsx ('a.xlsx',sheet = "public")
p <- na.omit(p)

hist(p$AEM,
     xlab = "AEM", 
     ylab = "Частота", 
     main = "AEM",
     col = "plum",
     freq = F 
)

mean_beaver6 <- mean(p$AEM)
sd_beaver6 <- sd(p$AEM)

curve(dnorm(x,
            mean = mean_beaver6, 
            sd = sd_beaver6),
      add = T 
)

hist(p$REM,
     xlab = "REM", 
     ylab = "Частота", 
     main = "REM",
     col = "plum", 
     freq = F 
)
mean_beaver7 <- mean(p$REM)
sd_beaver7 <- sd(p$REM)

curve(dnorm(x,
            mean = mean_beaver7, 
            sd = sd_beaver7),
      add = T 
)

hist(p$GROWTH,
     xlab = "GROWTH", 
     ylab = "Частота", 
     main = "GROWTH",
     col = "plum", 
     freq = F 
)
mean_beaver8 <- mean(p$GROWTH)
sd_beaver8 <- sd(p$GROWTH)

curve(dnorm(x,
            mean = mean_beaver8, 
            sd = sd_beaver8),
      add = T 
)

hist(p$ROAadj,
     xlab = "ROAadj", 
     ylab = "Частота", 
     main = "ROAadj",
     col = "plum", 
     freq = F 
)
mean_beaver9 <- mean(p$ROAadj)
sd_beaver9 <- sd(p$ROAadj)

curve(dnorm(x,
            mean = mean_beaver9, 
            sd = sd_beaver9),
      add = T 
)

hist(p$SIZE,
     xlab = "SIZE", 
     ylab = "Частота", 
     main = "SIZE",
     col = "plum", 
     freq = F 
)
mean_beaver10 <- mean(p$SIZE)
sd_beaver10 <- sd(p$SIZE)

curve(dnorm(x,
            mean = mean_beaver10, 
            sd = sd_beaver10),
      add = T 
)

hist(p$E_SCORE,
     xlab = "E_SCORE", 
     ylab = "Частота", 
     main = "E_SCORE",
     col = "plum", 
     freq = F 
)
mean_beaver11 <- mean(p$E_SCORE)
sd_beaver11 <- sd(p$E_SCORE)

curve(dnorm(x,
            mean = mean_beaver11, 
            sd = sd_beaver11),
      add = T 
)


hist(p$BOARDIN,
     xlab = "BOARDIN", 
     ylab = "Частота", 
     main = "BOARDIN",
     col = "plum", 
     freq = F 
)
mean_beaver12 <- mean(p$BOARDIN)
sd_beaver12 <- sd(p$BOARDIN)

curve(dnorm(x,
            mean = mean_beaver12, 
            sd = sd_beaver12),
      add = T 
)

hist(p$BOARDW,
     xlab = "BOARDW", 
     ylab = "Частота", 
     main = "BOARDW",
     col = "plum", 
     freq = F 
)
mean_beaver13 <- mean(p$BOARDW)
sd_beaver13 <- sd(p$BOARDW)

curve(dnorm(x,
            mean = mean_beaver13, 
            sd = sd_beaver13),
      add = T 
)

hist(p$BOARDM,
     xlab = "BOARDM", 
     ylab = "Частота", 
     main = "BOARDM",
     col = "plum", 
     freq = F 
)
mean_beaver14 <- mean(p$BOARDM)
sd_beaver14 <- sd(p$BOARDM)

curve(dnorm(x,
            mean = mean_beaver14, 
            sd = sd_beaver14),
      add = T 
)


#избавляемся от выбросов
p <- read.xlsx ('a.xlsx',sheet = "public")
p <- na.omit(p)

outliers <- boxplot(p$GROWTH, plot=FALSE)$out
ol <- p[-which(p$GROWTH %in% outliers),]
p<-ol

#логарифмируем BOARDM
p$BOARDMM <- log(p$BOARDM)

#проводим тесты на нормальность распределения
shapiro.test(p$AEM)
shapiro.test(p$REM)
shapiro.test(p$ROAadj)
shapiro.test(p$SIZE)
shapiro.test(p$GROWTH)
shapiro.test(p$E_SCORE)
shapiro.test(p$BOARDIN)
shapiro.test(p$BOARDMM)
shapiro.test(p$BOARDW)

#модель AEM pnnp - публичные и непубличные компании
q <- read.xlsx ('a.xlsx',sheet = "pnnp")
q <- na.omit(q)

outliers <- boxplot(q$REM, plot=FALSE)$out
ol <- q[-which(q$REM %in% outliers),]
q<-ol
outliers <- boxplot(q$GROWTH, plot=FALSE)$out
ol <- q[-which(q$GROWTH %in% outliers),]
q<-ol
#q$mark <- q$отрасль
#q$metal <- as.integer(grepl(pattern = "металл", q$отрасль))
#q$energy <- as.integer(grepl(pattern = "энергия", q$отрасль))

m.pooled4 <- plm(data = q, q$AEM ~ q$PUBLIC+q$E_SCORE+q$SIZE+q$ROAadj+q$GROWTH, model = "pooling")
summary(m.pooled4)

#проверка на гетероскедастичность, больше >0,05 - есть гомоскедастичность
bptest(m.pooled4)

#модель REM pnnp - публичные и непубличные компании
m.pooled5 <- plm(data = q, q$REM ~ q$PUBLIC+q$E_SCORE+q$SIZE+q$ROAadj+q$GROWTH, model = "pooling")
summary(m.pooled5)

#проверка на гетероскедастичность, больше >0,05 - есть гомоскедастичность
bptest(m.pooled5)

#провека на мультиколлинеарность
vif(m.pooled4)
vif(m.pooled5)

#вспомогательные модели для публичных и непубличных
m.p1 <- plm(data = q, q$AEM ~ q$PUBLIC+q$ECOMNG+q$SIZE+q$ROAadj+q$GROWTH, model = "pooling")
m.p2 <- plm(data = q, q$AEM ~ q$PUBLIC+q$ECOHARM+q$SIZE+q$ROAadj+q$GROWTH, model = "pooling")
m.p3 <- plm(data = q, q$AEM ~ q$PUBLIC+q$TRANSPAR+q$SIZE+q$ROAadj+q$GROWTH, model = "pooling")
m.p4 <- plm(data = q, q$REM ~ q$PUBLIC+q$ECOMNG+q$SIZE+q$ROAadj+q$GROWTH, model = "pooling")
m.p5 <- plm(data = q, q$REM ~ q$PUBLIC+q$ECOHARM+q$SIZE+q$ROAadj+q$GROWTH, model = "pooling")
m.p6 <- plm(data = q, q$REM ~ q$PUBLIC+q$TRANSPAR+q$SIZE+q$ROAadj+q$GROWTH, model = "pooling")

summary(m.p1)
summary(m.p2)
summary(m.p3)
summary(m.p4)
summary(m.p5)
summary(m.p6)

m.p7 <- plm(data = q, q$AB_CFOO ~ q$PUBLIC+q$E_SCORE+q$SIZE+q$ROAadj+q$GROWTH, model = "pooling")
m.p8 <- plm(data = q, q$AB_PROD ~ q$PUBLIC+q$E_SCORE+q$SIZE+q$ROAadj+q$GROWTH, model = "pooling")

summary(m.p7)
summary(m.p8)


#модели для публичных компаний
#избавляемся от выбросов
p <- read.xlsx ('a.xlsx',sheet = "public")
p <- na.omit(p)

outliers <- boxplot(p$GROWTH, plot=FALSE)$out
ol <- p[-which(p$GROWTH %in% outliers),]
p<-ol
#логарифмируем BOARDM
p$BOARDMM <- log(p$BOARDM)
#стоим базовую модель
m0 <- lm(data = p, p$REM ~ p$E_SCORE+p$SIZE+p$ROAadj+p$GROWTH+p$BOARDIN+p$BOARDMM+p$BOARDW)
#проверка на мультиколлинеарность
vif(m0)

#строим модели
p$mark <- p$отрасль
p$metal <- as.integer(grepl(pattern = "металл", p$отрасль))
p$energy <- as.integer(grepl(pattern = "энергия", p$отрасль))

m1 <- lm(data = p, p$AEM ~ p$E_SCORE+p$SIZE+p$ROAadj+p$GROWTH+p$BOARDIN+p$BOARDMM+p$BOARDW+p$metal+p$energy)
summary(m1)
m2 <- lm(data = p, p$REM ~ p$E_SCORE+p$SIZE+p$ROAadj+p$GROWTH+p$BOARDIN+p$BOARDMM+p$BOARDW+p$metal+p$energy)
summary(m2)

#проверка на гетероскедастичность
bptest(m1)
bptest(m2)

#корректируем коэффициенты
#если кластеры 
coeftest(m1, vcov=vcovHC(m1,type="HC0",cluster="CompanyName"))
coeftest(m2, vcov=vcovHC(m2,type="HC0",cluster="CompanyName"))

#если без кластеров
coeftest(m2, vcov. =vcovHC, type = 'HC0')
coeftest(m1, vcov. =vcovHC, type = 'HC0')

#попытка сделать распределение нормальным у AEM
p$AA <- log(p$AEM*100)

m1 <- lm(data = p, p$AA ~ p$E_SCORE+p$SIZE+p$ROAadj+p$GROWTH+p$BOARDIN+p$BOARDMM+p$BOARDW+p$metal+p$energy)
summary(m1)
m2 <- lm(data = p, p$REM ~ p$E_SCORE+p$SIZE+p$ROAadj+p$GROWTH+p$BOARDIN+p$BOARDMM+p$BOARDW+p$metal+p$energy)
summary(m2)

coeftest(m1, vcov=vcovHC(m1,type="HC0",cluster="CompanyName"))
coeftest(m2, vcov=vcovHC(m2,type="HC0",cluster="CompanyName"))
