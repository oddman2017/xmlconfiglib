xmlconfiglib
============

XML configuration for ini file profile style

Using Microsoft XML COM library to parse the xml file 

The file format is like below:


<?xml version="1.0" encoding="UTF-8"?>
<Securitylabel>
    <Object>
        <created value= "2010/08/24"/>
        <environment>
            <UID value="UID"/>
            <IP value="IP address"/>
            <MAC value="mac"/>
            <os_ver value="os_ver"/>
        </environment>
        <createdby value="scotter"/>
        <modified value="2010/8/24"/>
        <modifiedby value="somebody else"/>
        <title value="The document title"/>
        <UUID value="The document UUID number"/>
        <version value="document version new"/>
        <description value="the new description"/>
        <metadata value="select save control metadata"/>
        <signed value="file signature"/>
        <barcode value="bar code"/>
    </Object>
</Securitylabel>
