@echo off
path %path%;E:\Projects\Tools

set tbversion="964"
call makedll IBAPI 
set tbversion="27"
call makedll IBEnhancedAPI IBENHAPI
call makedll IBTWSSP
call makedll TBInfoBase
call makedll TickfileSP
