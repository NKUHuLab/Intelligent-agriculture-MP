library(xlsx)
library(readxl)
library(hydroGOF)
library(randomForest)
library(iml)
library(ggplot2)
library(circlize)
library(RColorBrewer)
library(dplyr)
library(randomForestExplainer)
library(pdp)
library(tcltk)
library(patchwork)
library(raster)
library(ggbreak)
library(reshape2)
library(grid)
library(ggpointdensity)
library(ggsci)
library(openxlsx)
library(writexl)
library(caret)
library(igraph)
library(rfPermute)
library(plspm)
library(ggalt)
library(ggpointdensity)
library(viridis)
library(hexbin)
library(cowplot)
library(mgcv)

setwd('E:/Phd/microplastic Soil/microplastic Soil')
MPs_threshold <- c(0.9391,0.5153,0.7757)
crop <- c('Maize','Rice','Wheat')
###read data
data_ML <- read.xlsx('Dataset/model data.xlsx', 1)
Test_predict_wb <- createWorkbook()
Test_predict_10fold_wb <- createWorkbook()
Train_predict_10fold_wb <- createWorkbook()
importance_predict_10fold_wb <- createWorkbook()
for (i in 1:length(crop)){
  addWorksheet(Test_predict_wb,sheetName = crop[i])
}
for (i in 1:length(crop)){
  addWorksheet(Test_predict_10fold_wb,sheetName = crop[i])
}
for (i in 1:length(crop)){
  addWorksheet(Train_predict_10fold_wb,sheetName = crop[i])
}
for (i in 1:length(crop)){
  addWorksheet(importance_predict_10fold_wb,sheetName = crop[i])
}
for (i in 1:length(crop)){
  cat("Iteration:", crop[i], "\n")
  data_crop <- data_ML[,c(1:21,(21+i),26:ncol(data_ML))]
  colnames(data_crop)[ncol(data_crop)]<-'index'
  data_crop$index<-as.numeric(data_crop$index)/MPs_threshold[i]
  dmy <- dummyVars(" ~ .", data = data_crop)
  data_crop <- data.frame(predict(dmy, newdata = data_crop))
  set.seed(123);disorder <- sample(nrow(data_crop),replace=F)
  fold_num <- floor(nrow(data_crop)/10)
  n_test <- data.frame()
  n_train <- data.frame()
  n_importance <- data.frame(MSX <- c(rep(0, times = (ncol(data_crop)-1))), 
                             NP<- c(rep(0, times = (ncol(data_crop)-1))))
  predict <- data.frame()
  for (k in 1:10){
    o <- disorder[(fold_num*(k-1)+1):(fold_num*k)]
    rf.data <- data_crop[-o,]
    rf <- randomForest(index~. , data = rf.data, 
                       ntree=1000 ,mtry=14,
                       proximity = F,
                       importance = T)
    p <- as.data.frame(predict(rf, data_crop[o,]))
    a <- as.data.frame(data_crop[o,ncol(data_crop)])
    p_train <- as.data.frame(predict(rf, rf.data))
    a_train <- as.data.frame(rf.data[,ncol(data_crop)])
    imp <- as.data.frame(rf$importance)
    r2 <- cor(p,a)
    R2 <- R2(p,a)
    rm <- Metrics::rmse(p[,1],a[,1])
    predict <- rbind(predict,cbind(a,p))
    n_test <- rbind(n_test,cbind(r2, R2, rm))
    #train
    r2_train <- cor(p_train,a_train)
    R2_train <- R2(p_train,a_train)
    rm_train <- Metrics::rmse(p_train[,1],a_train[,1])
    n_train <- rbind(n_train,cbind(r2_train, R2_train, rm_train))
    #importance
    n_importance <- imp + n_importance
    print(paste0(crop[i],' model iteration',k,': ',r2))
    print(paste0(crop[i],' model iteration',k,': ',R2))
  }
  colnames(predict) <- c('Observation','Prediction')
  colnames(n_test) <- c('r2','R2','RMSE')
  colnames(n_train) <- c('r2','R2','RMSE')
  n_importance <- n_importance/10
  n_importance[,3] <- rownames(imp)
  colnames(n_importance) <- c("MSE","Node","variables")
  writeData(Test_predict_wb, sheet = i, predict)
  writeData(Test_predict_10fold_wb, sheet = i, n_test)
  writeData(Train_predict_10fold_wb, sheet = i, n_train)
  writeData(importance_predict_10fold_wb, sheet = i, n_importance)
  print(paste0(crop[i],' R2:', mean(n_test$R2)))
}
saveWorkbook(Test_predict_wb, "paper/model build/Test_predict.xlsx", overwrite = TRUE)
saveWorkbook(Test_predict_10fold_wb, "paper/model build/Test_predict_10fold.xlsx", overwrite = TRUE)
saveWorkbook(Train_predict_10fold_wb, "paper/model build/Train_predict_10fold.xlsx", overwrite = TRUE)
saveWorkbook(importance_predict_10fold_wb, "paper/model build/Importance_10fold.xlsx", overwrite = TRUE)

##############Importance analysis#####################
data <- read.xlsx('Dataset/model data.xlsx', 1)
data_ML <- data[,c(9,38:72)]
data_ML <- data_ML[,-c(28,30,33,35)]
Train_predict_wb <- createWorkbook()
Importance_wb <- createWorkbook()
for (i in 1:length(crop)){
  addWorksheet(Train_predict_wb,sheetName = crop[i])
}
for (i in 1:length(crop)){
  addWorksheet(Importance_wb,sheetName = crop[i])
}
rf.list<-list()
for (i in 1:length(crop)){
  dataset <- data_ML[,c(1:21,(21+i),26:ncol(data_ML))]
  colnames(dataset)[22] <- 'Tillage'
  colnames(dataset)[ncol(dataset)]<-'index'
  dataset$index<-as.numeric(dataset$index)/MPs_threshold[i]
  dmy <- dummyVars(" ~ .", data = dataset)
  dataset <- data.frame(predict(dmy, newdata = dataset))
  
  rf.list[[crop[i]]]<-local({
    randomForest(index~. , data = dataset, 
                 ntree=1000,mtry=14,
                 proximity = T,
                 importance = T)})
  rf<-rf.list[[crop[i]]]
  p <- predict(rf, dataset)
  a <- dataset[,length(dataset[1,])]
  predict <- cbind(a,p)
  imp <- as.data.frame(rf$importance)
  imp[,3] <- rownames(imp)
  colnames(imp) <- c("MSE","Node","variables")
  
  writeData(Train_predict_wb, sheet = i, predict)
  writeData(Importance_wb, sheet = i, imp)
  print(i)
}
save(rf.list,file='paper/model build/Rda/rflist.rda')
saveWorkbook(Train_predict_wb, "paper/model build/Train_predict.xlsx", overwrite = TRUE)
saveWorkbook(Importance_wb, "paper/model build/Importance.xlsx", overwrite = TRUE)

load(file='paper/model build/Rda/rflist.rda')
md<-list()
mi<-list()
for (i in 1:length(crop)){
  dataset <- data_ML[,c(1:21,(21+i),26:ncol(data_ML))]
  colnames(dataset)[22] <- 'Tillage'
  colnames(dataset)[ncol(dataset)]<-'index'
  dataset$index<-as.numeric(dataset$index)/MPs_threshold[i]
  dmy <- dummyVars(" ~ .", data = dataset)
  dataset <- data.frame(predict(dmy, newdata = dataset))
  colindex <- c(colnames(dataset))
  colindexf <- factor(colindex[-length(colindex)],levels=colindex[-length(colindex)])
  
  rf<-rf.list[[i]]
  min_depth_frame<-min_depth_distribution(rf)
  md[[crop[i]]]<-min_depth_frame
  im_frame<-measure_importance(rf)
  im_frame[4]<-im_frame[4]/max(im_frame[4])
  im_frame[5]<-im_frame[5]/max(im_frame[5])
  mi[[crop[i]]]<-im_frame
  print(i)
}
save(md,mi,file='paper/model build/Rda/multi-importance.rda')

###multi-way plot
load(file='paper/model build/Rda/rflist.rda')
load(file='paper/model build/Rda/multi-importance.rda')
mdplot<-list()
miplot<-list()
for (i in 1:length(crop)){
  print(i)
  min_depth_frame<-md[[i]]
  mdplot[[crop[i]]]<-local({
    min_depth_frame=min_depth_frame
    plot_min_depth_distribution(min_depth_frame,k=15)+
      theme(axis.text=element_text(colour='black',size=10),
            axis.title=element_text(colour='black',size=15))+
      theme(panel.grid.major = element_blank(),panel.grid.minor = element_blank())+
      theme(legend.key.size = unit(0.5,'line'),legend.title = element_text(size=rel(0.6)),
            legend.text = element_text(size=rel(0.5)))
  })
  ggsave(paste0('paper/Plot/Importance/mp_',crop[i],'.pdf'),width=7,height=7)
  
  im_frame=mi[[i]]
  im_frame$p_value<-im_frame$p_value/5
  miplot[[crop[i]]]<-local({
    im_frame=im_frame
    plot_multi_way_importance(im_frame, x_measure = "mse_increase",
                              y_measure = "node_purity_increase",
                              size_measure = "p_value", no_of_labels = 15)+
      theme(axis.text=element_text(colour='black',size=10),
            axis.title=element_text(colour='black',size=15))+
      theme(axis.line=element_line(color='black'),
            axis.ticks.length=unit(0.5,'line'))+
      labs(x="MSE increase", y="Node Purity Increase")+
      theme(panel.grid.major = element_blank(),panel.grid.minor = element_blank())+
      coord_fixed(ratio=1)+
      theme(legend.position=c(0.1,0.8))
  })
  ggsave(paste0('paper/Plot/Importance/im_',crop[i],'.pdf'),width=5,height=5)
}
save(mdplot,miplot,file='paper/model build/Rda/importanceplot.rda')