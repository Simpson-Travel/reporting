install.packages("RGoogleAnalytics")
install.packages("RGA")
isntall.packages("lubridate")
install.packages('RDCOMClient', repos='http://www.omegahat.net/R')

library(RGoogleAnalytics)
client.id  <- "327238918544-fp6jmos9t658up3jde9i0old7crp08uf.apps.googleusercontent.com"
client.secret <- "eV2Cuhzs_502bxJsvgl13s_n"
token <- Auth(client.id,client.secret)
save(token,file="./token_file")

library(RGA)
authorize()


