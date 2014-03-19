ExcelUtils README
==========================
http://excelutils.sourceforge.net

rainsoft: http://www.try2it.com/blog
jokeway:  http://spaces.msn.com/members/jokeway

We're headache on making web report all long.

ExcelUtils is a helper to export excel report in java web project. 
It's like velocity, has own tags, but these tags is written in excel file. 
By these tags, you can custom your excel report format freely, 
not edit any your source, just ExcelUtils parses your excel template and fills
values to export your report.

It is based POI project and beanutils project. 
It uses excel and template language's profit to make web reports easily.

After my hardwork, the parser is finished finally, in which report is exported by Excel Template.
It's funtions include:

 1. ${model.name} means getting property of the name from the model object.
 2. ${!model.name} means that last cell and this cell merge if model.name value equals last cell value.
 3. #foreach model in ${list}��means that iterate list��modelId is implied index of the list.
 4. #each ${model} ${width1},${width2}��model can be a Map,JavaBean,Collection or Array object, #each key 
    will show all property of the model.${width?} means merge ${width?} cells. If only one 
    width, all property use the same width. If more than one, use the witdh in order, not set will use "1".
 5. ${list[0].name} means get the first object from list, then read the property of name.
 6. ${map(key)} get the value from the map by the key name.
 7. ${list[${index}].name} [] can be a variable.
 8. ${map(${key})} () can be a vriable.
 9. #sum qty on ${list} where name like/=str sum qty on ${list} collection by where condition.
10. In net.sf.excelutils.tags Package, you can implement ITag to exentd Tag key. eg, FooTag will parse #foo.
11. ExcelResult for webwork.
12. ${model${index}} support.
13. #call service.method("str", ${name}) call a method
14. #formual SUM(C${currentRowNo}:F${currentRowNo}) means output excel formula SUM(C?:F?) ? means currentRowNo. 

                                  
dependency library:

poi-3.5.jar (required)
commons-beanutils.jar (required)
commons-digester.jar (required)
commons-logging.jar (required)

ognl.jar (build required, webwork demo required)
xwork.jar (build required, webwork demo required)
webwork-2.1.7.jar (build required, webwork demo required)
oscore.jar (webwork deom required)

struts.jar (struts demo required)
