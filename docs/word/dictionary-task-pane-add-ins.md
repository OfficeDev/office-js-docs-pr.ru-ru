---
title: Создание надстройки области задач словаря
description: Узнайте, как создать надстройку области задач словаря
ms.date: 09/26/2019
localization_priority: Normal
ms.openlocfilehash: 2d79a40511d28cdf5d11c33435703009b1793dc2
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077227"
---
# <a name="create-a-dictionary-task-pane-add-in"></a><span data-ttu-id="a956f-103">Создание надстройки области задач словаря</span><span class="sxs-lookup"><span data-stu-id="a956f-103">Create a dictionary task pane add-in</span></span>


<span data-ttu-id="a956f-104">В этой статье представлены пример надстройки области задач и соответствующая веб-служба, которая предоставляет словарные статьи определений или синонимов из тезауруса к слову, выбранному пользователем в документе Word 2013.</span><span class="sxs-lookup"><span data-stu-id="a956f-104">This article shows you an example of a task pane add-in with an accompanying web service that provides dictionary definitions or thesaurus synonyms for the user's current selection in a Word 2013 document.</span></span> 

<span data-ttu-id="a956f-105">Надстройка словаря Office базируется на стандартной надстройке области задач с дополнительными функциональными возможностями поддержки запросов и отображения определений из словарной веб-службы XML в дополнительных расположениях пользовательского интерфейса приложения Office.</span><span class="sxs-lookup"><span data-stu-id="a956f-105">A dictionary Office Add-in is based on the standard task pane add-in with additional features to support querying and displaying definitions from a dictionary XML web service in additional places in the Office application's UI.</span></span> 

<span data-ttu-id="a956f-p101">В обычной надстройке области задач словаря пользователь выбирает слово или фразу в документе, после чего логика JavaScript надстройки передает выделенный фрагмент в XML-веб-службу поставщика словаря. Затем веб-страница этого поставщика обновляется, чтобы показать пользователю определения выделенного фрагмента. Компонент XML-веб-службы возвращает до трех определений в формате, определенном схемой XML OfficeDefinitions. Эти определения отображаются в ведущем приложении Office (в разных местах его пользовательского интерфейса). На рисунке 1 показано выделение фрагмента и отображение результатов при использовании надстройки словаря Bing, запущенной в Word 2013.</span><span class="sxs-lookup"><span data-stu-id="a956f-p101">In a typical dictionary task pane add-in, a user selects a word or phrase in their document, and the JavaScript logic behind the add-in passes this selection to the dictionary provider's XML web service. The dictionary provider's webpage then updates to show the definitions for the selection to the user. The XML web service component returns up to three definitions in the format defined by the OfficeDefinitions XML schema, which are then displayed to the user in other places in the hosting Office application's UI. Figure 1 shows the selection and display experience for a Bing-branded dictionary add-in that is running in Word 2013.</span></span>

<span data-ttu-id="a956f-110">*Рисунок 1. Надстройка словаря, отображающая определения выбранного слова*</span><span class="sxs-lookup"><span data-stu-id="a956f-110">*Figure 1. Dictionary add-in displaying definitions for the selected word*</span></span>

![Приложение Dictionary с определением.](../images/dictionary-agave-01.jpg)

<span data-ttu-id="a956f-112">Это зависит от вас, чтобы  определить, если щелкнув ссылку См. больше в HTML-интерфейсе надстройки словаря отображает дополнительные сведения в области задач или открывает отдельное окно браузера на полную веб-страницу для выбранного слова или фразы.</span><span class="sxs-lookup"><span data-stu-id="a956f-112">It is up to you to determine if clicking the **See More** link in the dictionary add-in's HTML UI displays more information within the task pane or opens a separate browser window to the full webpage for the selected word or phrase.</span></span>
<span data-ttu-id="a956f-113">На рисунке 2 показана команда контекстного меню **Define,** которая позволяет пользователям быстро запускать установленные словари.</span><span class="sxs-lookup"><span data-stu-id="a956f-113">Figure 2 shows the **Define** context menu command that enables users to quickly launch installed dictionaries.</span></span> <span data-ttu-id="a956f-114">На рисунках 3–5 перечислены все расположения в пользовательском интерфейсе Office, в которых словарные XML-службы предоставляют определения в Word 2013.</span><span class="sxs-lookup"><span data-stu-id="a956f-114">Figures 3 through 5 show the places in the Office UI where the dictionary XML services are used to provide definitions in Word 2013.</span></span>

<span data-ttu-id="a956f-115">*Рис. 2. Команда определения в контекстном меню*</span><span class="sxs-lookup"><span data-stu-id="a956f-115">*Figure 2. Define command in the context menu*</span></span>

![Определение контекстного меню.](../images/dictionary-agave-02.jpg)


<span data-ttu-id="a956f-117">*Рис. 3. Определения в областях проверки правописания*</span><span class="sxs-lookup"><span data-stu-id="a956f-117">*Figure 3. Definitions in the Spelling and Grammar panes*</span></span>

![Определения в области орфографии и грамматики.](../images/dictionary-agave-03.jpg)


<span data-ttu-id="a956f-119">*Рис. 4. Определения в области "Тезаурус"*</span><span class="sxs-lookup"><span data-stu-id="a956f-119">*Figure 4. Definitions in the Thesaurus pane*</span></span>

![Определения в области Тезаурус.](../images/dictionary-agave-04.jpg)


<span data-ttu-id="a956f-121">*Рис. 5. Определения в режиме чтения*</span><span class="sxs-lookup"><span data-stu-id="a956f-121">*Figure 5. Definitions in Reading Mode*</span></span>

![Определения в режиме чтения.](../images/dictionary-agave-05.jpg)

<span data-ttu-id="a956f-123&quot;>Чтобы создать надстройку области задач, которая выполняет поиск в словаре, необходимо создать два основных компонента:</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;a956f-123&quot;>To create a task pane add-in that provides a dictionary lookup, you create two main components:</span></span> 


- <span data-ttu-id=&quot;a956f-124&quot;>веб-службу XML, которая ищет определения в словарной службе, а затем возвращает результаты в формате XML, которые могут быть отображены в надстройке словаря;</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;a956f-124&quot;>An XML web service that looks up definitions from a dictionary service, and then returns those values in an XML format that can be consumed and displayed by the dictionary add-in.</span></span>
    
- <span data-ttu-id=&quot;a956f-125&quot;>надстройку области задач, которая отправляет выбранное пользователем слово или фразу в словарную веб-службу, отображает определения и может вставить эти значения в документ.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;a956f-125&quot;>A task pane add-in that submits the user's current selection to the dictionary web service, displays definitions, and can optionally insert those values into the document.</span></span>
    
<span data-ttu-id=&quot;a956f-126&quot;>В следующих разделах приведены примеры создания этих компонентов.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;a956f-126&quot;>The following sections provide examples of how to create these components.</span></span>

## <a name=&quot;creating-a-dictionary-xml-web-service&quot;></a><span data-ttu-id=&quot;a956f-127&quot;>Создание словарной веб-службы XML</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;a956f-127&quot;>Creating a dictionary XML web service</span></span>


<span data-ttu-id=&quot;a956f-p103&quot;>Веб-служба XML должна возвращать запросы веб-служб в виде XML-кода, который соответствует XML-схеме OfficeDefinitions. В двух следующих разделах описывается XML-схема OfficeDefinitions и предоставлен пример возможности кодирования веб-службы XML, возвращающей запросы в этом формате XML.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;a956f-p103&quot;>The XML web service must return queries to the web service as XML that conforms to the OfficeDefinitions XML schema. The following two sections describe the OfficeDefinitions XML schema, and provide an example of how to code an XML web service that returns queries in that XML format.</span></span>


### <a name=&quot;officedefinitions-xml-schema&quot;></a><span data-ttu-id=&quot;a956f-130&quot;>XML-схема OfficeDefinitions</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;a956f-130&quot;>OfficeDefinitions XML schema</span></span>

<span data-ttu-id=&quot;a956f-131&quot;>В следующем коде отображается XSD для XML-схемы OfficeDefinitions.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;a956f-131&quot;>The following code shows the XSD for the OfficeDefinitions XML Schema.</span></span>


```XML
<?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?>
<xs:schema
  xmlns:xsi=&quot;http://www.w3.org/2001/XMLSchema-instance&quot;
  xmlns:xs=&quot;https://www.w3.org/2001/XMLSchema&quot;
  targetNamespace=&quot;http://schemas.microsoft.com/NLG/2011/OfficeDefinitions&quot;
  xmlns=&quot;http://schemas.microsoft.com/NLG/2011/OfficeDefinitions&quot;>
  <xs:element name=&quot;Result&quot;>
    <xs:complexType>
      <xs:sequence>
        <xs:element name=&quot;SeeMoreURL&quot; type=&quot;xs:anyURI&quot;/>
        <xs:element name=&quot;Definitions&quot; type=&quot;DefinitionListType&quot;/>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
  <xs:complexType name=&quot;DefinitionListType&quot;>
    <xs:sequence>
      <xs:element name=&quot;Definition&quot; maxOccurs=&quot;3&quot;>
        <xs:simpleType>
          <xs:restriction base=&quot;xs:normalizedString&quot;>
            <xs:maxLength value=&quot;400&quot;/>
          </xs:restriction>
        </xs:simpleType>
      </xs:element>
    </xs:sequence>
  </xs:complexType>
</xs:schema>
```

<span data-ttu-id=&quot;a956f-132&quot;>Возвращенный XML, соответствующий схеме OfficeDefinitions, состоит из корневого элемента, который содержит элемент с от нуля до трех детских элементов, каждый из которых содержит определения длиной не более `Result` `Definitions` `Definition` 400 символов.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;a956f-132&quot;>Returned XML that conforms to the OfficeDefinitions schema consists of a root `Result` element that contains a `Definitions` element with from zero to three `Definition` child elements, each of which contains definitions that are no more than 400 characters in length.</span></span> <span data-ttu-id=&quot;a956f-133&quot;>Кроме того, в элементе должен быть указан URL-адрес полной страницы на сайте `SeeMoreURL` словаря.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;a956f-133&quot;>Additionally, the URL to the full page on the dictionary site must be provided in the `SeeMoreURL` element.</span></span> <span data-ttu-id=&quot;a956f-134&quot;>В следующем примере показана структура возвращенного XML-кода, соответствующего схеме OfficeDefinitions.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;a956f-134&quot;>The following example shows the structure of returned XML that conforms to the OfficeDefinitions schema.</span></span>

```XML
<?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?>
<Result xmlns=&quot;http://schemas.microsoft.com/NLG/2011/OfficeDefinitions&quot;>
  <SeeMoreURL xmlns=&quot;&quot;>www.bing.com/dictionary/search?q=example</SeeMoreURL>
  <Definitions xmlns=&quot;&quot;>
    <Definition>Definition1</Definition>
    <Definition>Definition2</Definition>
    <Definition>Definition3</Definition>
  </Definitions>
 </Result>

```


### <a name=&quot;sample-dictionary-xml-web-service&quot;></a><span data-ttu-id=&quot;a956f-135&quot;>Пример словарной веб-службы XML</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;a956f-135&quot;>Sample dictionary XML web service</span></span>

<span data-ttu-id=&quot;a956f-136&quot;>Приведенный ниже код C# предоставляет простой пример написания кода для веб-службы XML, которая возвращает результат запроса словаря в XML-формате OfficeDefinitions.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;a956f-136&quot;>The following C# code provides a simple example of how to write code for an XML web service that returns the result of a dictionary query in the OfficeDefinitions XML format.</span></span>


```cs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Xml;
using System.Text;
using System.IO;
using System.Net;

/// <summary>
/// Summary description for _Default
/// </summary>
[WebService(Namespace = &quot;http://tempuri.org/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// To allow this web service to be called from script, using ASP.NET AJAX, uncomment the following line. 
// [System.Web.Script.Services.ScriptService]
public class WebService : System.Web.Services.WebService {

    public WebService () {

        // Uncomment the following line if using designed components 
        // InitializeComponent(); 
    }

    // You can replace this method entirely with your own method that gets definitions
    // from your data source, and then formats it into the OfficeDefinitions XML format. 
    // If you need a reference for constructing the returned XML, you can use this example as a basis.
    [WebMethod]
    public XmlDocument Define(string word)
    {

        StringBuilder sb = new StringBuilder();
        XmlWriter writer = XmlWriter.Create(sb);
        {
            writer.WriteStartDocument();
            
                writer.WriteStartElement("Result", "http://schemas.microsoft.com/NLG/2011/OfficeDefinitions");

            // See More URL should be changed to the dictionary publisher's page for that word on their website.
                    writer.WriteElementString("SeeMoreURL", "http://www.bing.com/search?q=" + word);

                    writer.WriteStartElement("Definitions");
            
                        writer.WriteElementString("Definition", "Definition 1 of " + word);
                        writer.WriteElementString("Definition", "Definition 2 of " + word);
                        writer.WriteElementString("Definition", "Definition 3 of " + word);
                   
                    writer.WriteEndElement();


                writer.WriteEndElement();
            
            writer.WriteEndDocument();
        }
        writer.Close();

        XmlDocument doc = new XmlDocument();
        doc.LoadXml(sb.ToString());

        return doc;
    }
}
```


## <a name="creating-the-components-of-a-dictionary-add-in"></a><span data-ttu-id="a956f-137">Создание компонентов надстройки словаря</span><span class="sxs-lookup"><span data-stu-id="a956f-137">Creating the components of a dictionary add-in</span></span>


<span data-ttu-id="a956f-138">Надстройка словаря состоит из трех основных файлов компонентов.</span><span class="sxs-lookup"><span data-stu-id="a956f-138">A dictionary add-in consists of three main component files:</span></span>


- <span data-ttu-id="a956f-139">XML-файл манифеста, который описывает надстройку.</span><span class="sxs-lookup"><span data-stu-id="a956f-139">An XML manifest file that describes the add-in.</span></span>
    
- <span data-ttu-id="a956f-140">HTML-файл, который предоставляет пользовательский интерфейс надстройки.</span><span class="sxs-lookup"><span data-stu-id="a956f-140">An HTML file that provides the add-in's UI.</span></span>
    
- <span data-ttu-id="a956f-141">Файл JavaScript, который содержит логику для получения выделенного пользователем фрагмента из документа, отправки выбранного слова или фразы в веб-службу и отображения возвращенных результатов в пользовательском интерфейсе надстройки.</span><span class="sxs-lookup"><span data-stu-id="a956f-141">A JavaScript file that provides logic to get the user's selection from the document, sends the selection as a query to the web service, and then displays returned results in the add-in's UI.</span></span>
    

### <a name="creating-a-dictionary-add-ins-manifest-file"></a><span data-ttu-id="a956f-142">Создание файла манифеста надстройки словаря</span><span class="sxs-lookup"><span data-stu-id="a956f-142">Creating a dictionary add-in's manifest file</span></span>

<span data-ttu-id="a956f-143">Ниже приведен пример файла манифеста для надстройки словаря.</span><span class="sxs-lookup"><span data-stu-id="a956f-143">The following is an example manifest file for a dictionary add-in.</span></span>


```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <Id>7164e750-dc86-49c0-b548-1bac57abdc7c</Id>
  <Version>15.0</Version>
  <ProviderName>Microsoft Office Demo Dictionary</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <!--DisplayName is the name that will appear in the user's list of applications.-->
  <DisplayName DefaultValue="Microsoft Office Demo Dictionary" />
  <!--Description is a 2-3 sentence description of this dictionary. -->
  <Description DefaultValue="The Microsoft Office Demo Dictionary is an example built to demonstrate how a publisher could create a dictionary that integrates with Office. It does not return real definitions." />
  <!--IconUrl is the URI for the icon that will appear in the user's list of applications.-->
  <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <!--Capabilities specifies the kind of Office application your dictionary add-in will support. You shouldn't have to modify this area.-->
  <Capabilities>
    <Capability Name="Workbook"/>
    <Capability Name="Document"/>
    <Capability Name="Project"/>
  </Capabilities>
  <DefaultSettings>
    <!--SourceLocation is the URL for your dictionary-->
    <SourceLocation DefaultValue="http://christophernlg/ExampleDictionary/DictionaryHome.html" />
  </DefaultSettings>
  <!--Permissions is the set of permissions a user will have to give your dictionary. If you need write access, such as to allow a user to replace the highlighted word with a synonym, use ReadWriteDocument. -->
  <Permissions>ReadDocument</Permissions>
  <Dictionary>
    <!--TargetDialects is the set of regional languages your dictionary contains. For example, if your dictionary applies to Spanish (Mexico) and Spanish (Peru), but not Spanish (Spain), you can specify that here. Do not put more than one language (for example, Spanish and English) here. Publish separate languages as separate dictionaries. -->
    <TargetDialects>
      <TargetDialect>EN-AU</TargetDialect>
      <TargetDialect>EN-BZ</TargetDialect>
      <TargetDialect>EN-CA</TargetDialect>
      <TargetDialect>EN-029</TargetDialect>
      <TargetDialect>EN-HK</TargetDialect>
      <TargetDialect>EN-IN</TargetDialect>
      <TargetDialect>EN-ID</TargetDialect>
      <TargetDialect>EN-IE</TargetDialect>
      <TargetDialect>EN-JM</TargetDialect>
      <TargetDialect>EN-MY</TargetDialect>
      <TargetDialect>EN-NZ</TargetDialect>
      <TargetDialect>EN-PH</TargetDialect>
      <TargetDialect>EN-SG</TargetDialect>
      <TargetDialect>EN-ZA</TargetDialect>
      <TargetDialect>EN-TT</TargetDialect>
      <TargetDialect>EN-GB</TargetDialect>
      <TargetDialect>EN-US</TargetDialect>
      <TargetDialect>EN-ZW</TargetDialect>
    </TargetDialects>
    <!--QueryUri is the address of this dictionary's XML web service (which is used to put definitions in additional contexts, such as the spelling checker.)-->
    <QueryUri DefaultValue="http://christophernlg/ExampleDictionary/WebService.asmx/Define?word="/>
    <!--Citation Text, Dictionary Name, and Dictionary Home Page will be combined to form the citation line (for example, this would produce "Examples by: Microsoft", where "Microsoft" is a hyperlink to http://www.microsoft.com).-->
    <CitationText DefaultValue="Examples by: " />
    <DictionaryName DefaultValue="Microsoft" />
    <DictionaryHomePage DefaultValue="http://www.microsoft.com" />
  </Dictionary>
</OfficeApp>
```

<span data-ttu-id="a956f-144">Элемент и его детские элементы, специфические для создания файла манифеста надстройки словаря, описаны `Dictionary` в следующих разделах.</span><span class="sxs-lookup"><span data-stu-id="a956f-144">The `Dictionary` element and its child elements that are specific to creating a dictionary add-in's manifest file are described in the following sections.</span></span> <span data-ttu-id="a956f-145">Сведения о других элементах в файле манифеста см. в Office [XML-манифеста надстройки.](../develop/add-in-manifests.md)</span><span class="sxs-lookup"><span data-stu-id="a956f-145">For information about the other elements in the manifest file, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>


### <a name="dictionary-element"></a><span data-ttu-id="a956f-146">Элемент Dictionary</span><span class="sxs-lookup"><span data-stu-id="a956f-146">Dictionary element</span></span>


<span data-ttu-id="a956f-147">Определяет параметры надстроек словаря.</span><span class="sxs-lookup"><span data-stu-id="a956f-147">Specifies settings for dictionary add-ins.</span></span>

 <span data-ttu-id="a956f-148">**Родительский элемент**</span><span class="sxs-lookup"><span data-stu-id="a956f-148">**Parent element**</span></span>

 `<OfficeApp>`

 <span data-ttu-id="a956f-149">**Дочерние элементы**</span><span class="sxs-lookup"><span data-stu-id="a956f-149">**Child elements**</span></span>

 <span data-ttu-id="a956f-150">`<TargetDialects>`, `<QueryUri>`, `<CitationText>`, `<DictionaryName>`, `<DictionaryHomePage>`</span><span class="sxs-lookup"><span data-stu-id="a956f-150">`<TargetDialects>`, `<QueryUri>`, `<CitationText>`, `<DictionaryName>`, `<DictionaryHomePage>`</span></span>

 <span data-ttu-id="a956f-151">**Замечания**</span><span class="sxs-lookup"><span data-stu-id="a956f-151">**Remarks**</span></span>

<span data-ttu-id="a956f-152">Элемент и его детские элементы добавляются в манифест надстройки области задач при создании надстройки `Dictionary` словаря.</span><span class="sxs-lookup"><span data-stu-id="a956f-152">The `Dictionary` element and its child elements are added to the manifest of a task pane add-in when you create a dictionary add-in.</span></span>


#### <a name="targetdialects-element"></a><span data-ttu-id="a956f-153">Элемент TargetDialects</span><span class="sxs-lookup"><span data-stu-id="a956f-153">TargetDialects element</span></span>


<span data-ttu-id="a956f-p106">Определяет региональные языки, которые поддерживает этот словарь. Обязательный для надстроек словаря.</span><span class="sxs-lookup"><span data-stu-id="a956f-p106">Specifies the regional languages that this dictionary supports. Required for dictionary add-ins.</span></span>

 <span data-ttu-id="a956f-156">**Родительский элемент**</span><span class="sxs-lookup"><span data-stu-id="a956f-156">**Parent element**</span></span>

 `<Dictionary>`

 <span data-ttu-id="a956f-157">**Дочерний элемент**</span><span class="sxs-lookup"><span data-stu-id="a956f-157">**Child element**</span></span>

 `<TargetDialect>`

 <span data-ttu-id="a956f-158">**Замечания**</span><span class="sxs-lookup"><span data-stu-id="a956f-158">**Remarks**</span></span>

<span data-ttu-id="a956f-159">Элемент и его детские элементы указывают набор региональных `TargetDialects` языков, которые содержатся в словаре.</span><span class="sxs-lookup"><span data-stu-id="a956f-159">The `TargetDialects` element and its child elements specify the set of regional languages your dictionary contains.</span></span> <span data-ttu-id="a956f-160">Например, если словарь применяется к испанскому языку, на котором разговаривают в Мексике и Перу, но не в Испании, это можно указать в данном элементе.</span><span class="sxs-lookup"><span data-stu-id="a956f-160">For example, if your dictionary applies to both Spanish (Mexico) and Spanish (Peru), but not Spanish (Spain), you can specify that in this element.</span></span> <span data-ttu-id="a956f-161">Не указывайте в этом манифесте более одного языка (например, испанский и английский).</span><span class="sxs-lookup"><span data-stu-id="a956f-161">Do not specify more than one language (e.g., Spanish and English) in this manifest.</span></span> <span data-ttu-id="a956f-162">Публикуйте разные языки для отдельных словарей.</span><span class="sxs-lookup"><span data-stu-id="a956f-162">Publish separate languages as separate dictionaries.</span></span>

 <span data-ttu-id="a956f-163">**Пример**</span><span class="sxs-lookup"><span data-stu-id="a956f-163">**Example**</span></span>

```XML
<TargetDialects>
  <TargetDialect>EN-AU</TargetDialect>
  <TargetDialect>EN-BZ</TargetDialect>
  <TargetDialect>EN-CA</TargetDialect>
  <TargetDialect>EN-029</TargetDialect>
  <TargetDialect>EN-HK</TargetDialect>
  <TargetDialect>EN-IN</TargetDialect>
  <TargetDialect>EN-ID</TargetDialect>
  <TargetDialect>EN-IE</TargetDialect>
  <TargetDialect>EN-JM</TargetDialect>
  <TargetDialect>EN-MY</TargetDialect>
  <TargetDialect>EN-NZ</TargetDialect>
  <TargetDialect>EN-PH</TargetDialect>
  <TargetDialect>EN-SG</TargetDialect>
  <TargetDialect>EN-ZA</TargetDialect>
  <TargetDialect>EN-TT</TargetDialect>
  <TargetDialect>EN-GB</TargetDialect>
  <TargetDialect>EN-US</TargetDialect>
  <TargetDialect>EN-ZW</TargetDialect>
</TargetDialects>
```


#### <a name="targetdialect-element"></a><span data-ttu-id="a956f-164">Элемент TargetDialect</span><span class="sxs-lookup"><span data-stu-id="a956f-164">TargetDialect element</span></span>


<span data-ttu-id="a956f-p108">Определяет региональный язык, который поддерживает этот словарь. Обязательный для надстроек словаря.</span><span class="sxs-lookup"><span data-stu-id="a956f-p108">Specifies a regional language that this dictionary supports. Required for dictionary add-ins.</span></span>

 <span data-ttu-id="a956f-167">**Родительский элемент**</span><span class="sxs-lookup"><span data-stu-id="a956f-167">**Parent element**</span></span>

 `<TargetDialects>`

 <span data-ttu-id="a956f-168">**Примечания**</span><span class="sxs-lookup"><span data-stu-id="a956f-168">**Remarks**</span></span>

<span data-ttu-id="a956f-169">Укажите значение регионального языка в формате тегов `language` RFC1766, например EN-US.</span><span class="sxs-lookup"><span data-stu-id="a956f-169">Specify the value for a regional language in the RFC1766  `language` tag format, such as EN-US.</span></span>

 <span data-ttu-id="a956f-170">**Пример**</span><span class="sxs-lookup"><span data-stu-id="a956f-170">**Example**</span></span>


```XML
<TargetDialect>EN-US</TargetDialect>
```


#### <a name="queryuri-element"></a><span data-ttu-id="a956f-171">Элемент QueryUri</span><span class="sxs-lookup"><span data-stu-id="a956f-171">QueryUri element</span></span>


<span data-ttu-id="a956f-p109">Определяет конечную точку службы запросов словаря. Обязательный элемент для надстроек словаря.</span><span class="sxs-lookup"><span data-stu-id="a956f-p109">Specifies the endpoint for the dictionary query service. Required for dictionary add-ins.</span></span>

 <span data-ttu-id="a956f-174">**Родительский элемент**</span><span class="sxs-lookup"><span data-stu-id="a956f-174">**Parent element**</span></span>

 `<Dictionary>`

 <span data-ttu-id="a956f-175">**Замечания**</span><span class="sxs-lookup"><span data-stu-id="a956f-175">**Remarks**</span></span>

<span data-ttu-id="a956f-p110">Это универсальный код ресурса (URI) XML-веб-службы поставщика словаря. К этому URI добавляется строка запроса с надлежащими escape-символами.</span><span class="sxs-lookup"><span data-stu-id="a956f-p110">This is the URI of the XML web service for the dictionary provider. The properly escaped query will be appended to this URI.</span></span> 

 <span data-ttu-id="a956f-178">**Пример**</span><span class="sxs-lookup"><span data-stu-id="a956f-178">**Example**</span></span>


```XML
<QueryUri DefaultValue="http://msranlc-lingo1/proof.aspx?q="/>
```


#### <a name="citationtext-element"></a><span data-ttu-id="a956f-179">Элемент CitationText</span><span class="sxs-lookup"><span data-stu-id="a956f-179">CitationText element</span></span>


<span data-ttu-id="a956f-p111">Определяет текст, который будет использоваться в ссылках. Обязательный элемент для надстроек словаря.</span><span class="sxs-lookup"><span data-stu-id="a956f-p111">Specifies the text to use in citations. Required for dictionary add-ins.</span></span>

 <span data-ttu-id="a956f-182">**Родительский элемент**</span><span class="sxs-lookup"><span data-stu-id="a956f-182">**Parent element**</span></span>

 `<Dictionary>`

 <span data-ttu-id="a956f-183">**Замечания**</span><span class="sxs-lookup"><span data-stu-id="a956f-183">**Remarks**</span></span>

<span data-ttu-id="a956f-184">В этом элементе указывается начальный текст ссылки, который будет отображаться в строке под контентом, возвращенным из веб-службы (например, "Источник:" или "Предоставлено:").</span><span class="sxs-lookup"><span data-stu-id="a956f-184">This element specifies the beginning of the citation text that will be displayed on a line below the content that is returned from the web service (for example, "Results by: " or "Powered by: ").</span></span>

<span data-ttu-id="a956f-185">Для этого элемента можно указать значения для дополнительных локализов с помощью `Override` элемента.</span><span class="sxs-lookup"><span data-stu-id="a956f-185">For this element, you can specify values for additional locales by using the `Override` element.</span></span> <span data-ttu-id="a956f-186">Например, если пользователь использует версию Office на испанском языке, но задействует английский словарь, то в строке ссылки будет написано "Resultados por: Bing", а не "Results by: Bing" или "Источник: Bing".</span><span class="sxs-lookup"><span data-stu-id="a956f-186">For example, if a user is running the Spanish SKU of Office, but using an English dictionary, this allows the citation line to read "Resultados por: Bing" rather than "Results by: Bing".</span></span> <span data-ttu-id="a956f-187">Дополнительные сведения о том, как указать значения для дополнительных локализов, см. в разделе "Предоставление параметров для различных локализов" в манифесте XML надстройки Office [надстройки.](../develop/add-in-manifests.md)</span><span class="sxs-lookup"><span data-stu-id="a956f-187">For more information about how to specify values for additional locales, see the section "Providing settings for different locales" in [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>

 <span data-ttu-id="a956f-188">**Пример**</span><span class="sxs-lookup"><span data-stu-id="a956f-188">**Example**</span></span>


```XML
<CitationText DefaultValue="Results by: " />
```


#### <a name="dictionaryname-element"></a><span data-ttu-id="a956f-189">Элемент DictionaryName</span><span class="sxs-lookup"><span data-stu-id="a956f-189">DictionaryName element</span></span>


<span data-ttu-id="a956f-p113">Определяет имя этого словаря. Обязательный элемент для надстроек словаря.</span><span class="sxs-lookup"><span data-stu-id="a956f-p113">Specifies the name of this dictionary. Required for dictionary add-ins.</span></span>

 <span data-ttu-id="a956f-192">**Родительский элемент**</span><span class="sxs-lookup"><span data-stu-id="a956f-192">**Parent element**</span></span>

 `<Dictionary>`

 <span data-ttu-id="a956f-193">**Замечания**</span><span class="sxs-lookup"><span data-stu-id="a956f-193">**Remarks**</span></span>

<span data-ttu-id="a956f-p114">В этом элементе указывается текст ссылки на источник. Текст ссылки на источник отображается в строчке под контентом, возвращенным веб-службой.</span><span class="sxs-lookup"><span data-stu-id="a956f-p114">This element specifies the link text in the citation text. Citation text is displayed on a line below the content that is returned from the web service.</span></span>

<span data-ttu-id="a956f-196">В этом элементе можно задать значения для дополнительных языковых стандартов.</span><span class="sxs-lookup"><span data-stu-id="a956f-196">For this element, you can specify values for additional locales.</span></span>

 <span data-ttu-id="a956f-197">**Пример**</span><span class="sxs-lookup"><span data-stu-id="a956f-197">**Example**</span></span>

```XML
<DictionaryName DefaultValue="Bing Dictionary" />
```


#### <a name="dictionaryhomepage-element"></a><span data-ttu-id="a956f-198">Элемент DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="a956f-198">DictionaryHomePage element</span></span>


<span data-ttu-id="a956f-p115">Определяет URL-адрес домашней страницы словаря. Обязательный элемент для надстроек словаря.</span><span class="sxs-lookup"><span data-stu-id="a956f-p115">Specifies the URL of the home page for the dictionary. Required for dictionary add-ins.</span></span>

 <span data-ttu-id="a956f-201">**Родительский элемент**</span><span class="sxs-lookup"><span data-stu-id="a956f-201">**Parent element**</span></span>

 `<Dictionary>`

 <span data-ttu-id="a956f-202">**Замечания**</span><span class="sxs-lookup"><span data-stu-id="a956f-202">**Remarks**</span></span>

<span data-ttu-id="a956f-p116">В этом элементе указывается URL-адрес источника. Текст ссылки на источник отображается в строчке под контентом, возвращенным веб-службой.</span><span class="sxs-lookup"><span data-stu-id="a956f-p116">This element specifies the link URL in the citation text. Citation text is displayed on a line below the content that is returned from the web service.</span></span>

<span data-ttu-id="a956f-205">В этом элементе можно задать значения для дополнительных языковых стандартов.</span><span class="sxs-lookup"><span data-stu-id="a956f-205">For this element, you can specify values for additional locales.</span></span>

 <span data-ttu-id="a956f-206">**Пример**</span><span class="sxs-lookup"><span data-stu-id="a956f-206">**Example**</span></span>


```XML
<DictionaryHomePage DefaultValue="http://www.bing.com" />
```


### <a name="creating-a-dictionary-add-ins-html-user-interface"></a><span data-ttu-id="a956f-207">Создание пользовательского интерфейса HTML для надстройки словаря</span><span class="sxs-lookup"><span data-stu-id="a956f-207">Creating a dictionary add-in's HTML user interface</span></span>

<span data-ttu-id="a956f-p117">В двух следующих примерах показаны HTML- и CSS-файлы для пользовательского интерфейса демонстрационной надстройки словаря. Чтобы просмотреть, как отображается пользовательский интерфейс в надстройке области задач, изучите рис. 6, который приведен сразу после кода. Чтобы узнать, как реализация JavaScript в файле Dictionary.js предоставляет логику программирования для этого пользовательского интерфейса HTML, см. раздел "Составление реализации JavaScript" ниже.</span><span class="sxs-lookup"><span data-stu-id="a956f-p117">The following two examples show the HTML and CSS files for the UI of the Demo Dictionary add-in. To view how the UI is displayed in the add-in's task pane, see Figure 6 following the code. To see how the implementation of the JavaScript in the Dictionary.js file provides programming logic for this HTML UI, see "Writing the JavaScript implementation" immediately following this section.</span></span>

```HTML
<!DOCTYPE html>
<html>

<head>
<meta http-equiv="X-UA-Compatible" content="IE=Edge"/>

<!--The title will not be shown but is supplied to ensure valid HTML.-->
<title>Example Dictionary</title>

<!--Required library includes.-->
<script type="text/javascript" src="http://ajax.microsoft.com/ajax/4.0/1/MicrosoftAjax.js"></script>
<script type="text/javascript" src="office.js"></script>

<!--Optional library includes.-->
<script type="text/javascript" src="http://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.5.1.js"></script>

<!--App-specific CSS and JS.-->
<link rel="Stylesheet" type="text/css" href="style.css" />
<script type="text/ecmascript" src="dictionary.js"></script>
</head>

<body>
<div id="mainContainer">
    <div id="header">
        <span id="headword"></span>
        <span id="pronunciation">(<a id="pronunciationLink">Pronounce</a>)</span>
    </div>
    <ol id="definitions">
    </ol>
    <div id="SeeMore">
    <a id="SeeMoreLink">See More...</a>
    </div>
</div>
</body>

</html>
```

<span data-ttu-id="a956f-211">В приведенном ниже примере показано содержание Style.css.</span><span class="sxs-lookup"><span data-stu-id="a956f-211">The following example shows the contents of Style.css.</span></span>

```CSS
#mainContainer
{
    font-family: Segoe UI;
    font-size: 11pt;
}

#headword
{
    font-family: Segoe UI Semibold;
    color: #262626;
}

#pronunciation
{
    margin-left: 2px;
    margin-right: 2px;
}

#definitions
{
    font-size: 8.5pt;
}
a
{
    font-size: 8pt;
    color: #336699;
    text-decoration: none;
}
a:visited
{
    color: #993366;
}
a:hover, a:active
{  
    text-decoration: underline;
}
```

<span data-ttu-id="a956f-212">*Рис. 6. Демонстрационный пользовательский интерфейс словаря*</span><span class="sxs-lookup"><span data-stu-id="a956f-212">*Figure 6. Demo dictionary UI*</span></span>

![Пользовательский интерфейс демо-словаря.](../images/dictionary-agave-06.jpg)


### <a name="writing-the-javascript-implementation"></a><span data-ttu-id="a956f-214">Реализация JavaScript</span><span class="sxs-lookup"><span data-stu-id="a956f-214">Writing the JavaScript implementation</span></span>


<span data-ttu-id="a956f-p118">В приведенном ниже примере показана реализация JavaScript в файле Dictionary.js, которая вызывается с HTML-страницы надстройки и предоставляет программную логику для надстройки Demo Dictionary. В этом сценарии используется вышеописанная XML-веб-служба. Если поместить сценарий в тот же каталог, что и пример веб-службы, он будет получать определения из этой службы. Его можно использовать с общедоступной XML-веб-службой, соответствующей схеме OfficeDefinitions. Для этого измените переменную `xmlServiceURL` в начале файла, а затем замените ключ API Bing для произношений на правильно зарегистрированный.</span><span class="sxs-lookup"><span data-stu-id="a956f-p118">The following example shows the JavaScript implementation in the Dictionary.js file that is called from the add-in's HTML page to provide the programming logic for the Demo Dictionary add-in. This script reuses the XML web service described previously. When placed in the same directory as the example web service, the script will get definitions from that service. It can be used with a public OfficeDefinitions-conforming XML web service by modifying the  `xmlServiceURL` variable at the top of the file, and then replacing the Bing API key for pronunciations with a properly registered one.</span></span>

<span data-ttu-id="a956f-219">Основные участники API Office JavaScript (Office.js), которые вызваны из этой реализации, являются следующими:</span><span class="sxs-lookup"><span data-stu-id="a956f-219">The primary members of the Office JavaScript API (Office.js) that are called from this implementation are as follows:</span></span>


- <span data-ttu-id="a956f-220">Событие [](/javascript/api/office) инициализации объекта, которое повышается при инициализации контекста надстройки, и предоставляет доступ к экземпляру объекта Document, который представляет документ, с которым взаимодействует надстройка. `Office` [](/javascript/api/office/office.document)</span><span class="sxs-lookup"><span data-stu-id="a956f-220">The [initialize](/javascript/api/office) event of the `Office` object, which is raised when the add-in context is initialized, and provides access to a [Document](/javascript/api/office/office.document) object instance that represents the document the add-in is interacting with.</span></span>
    
- <span data-ttu-id="a956f-221">Метод [addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) объекта, который вызван в функцию, чтобы добавить обработник событий для события `Document` `initialize` [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) документа для прослушивания изменений выбора пользователя.</span><span class="sxs-lookup"><span data-stu-id="a956f-221">The [addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) method of the `Document` object, which is called in the `initialize` function to add an event handler for the [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) event of the document to listen for user selection changes.</span></span>
    
- <span data-ttu-id="a956f-222">Метод [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) объекта, который вызывается в функции при поднятии обработника событий, чтобы получить выбранное пользователем слово или фразу, принудить его к простому тексту, а затем выполнить функцию асинхронного `Document` `tryUpdatingSelectedWord()` `SelectionChanged` `selectedTextCallback` вызова.</span><span class="sxs-lookup"><span data-stu-id="a956f-222">The [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) method of the `Document` object, which is called in the `tryUpdatingSelectedWord()` function when the `SelectionChanged` event handler is raised to get the word or phrase the user selected, coerce it to plain text, and then execute the `selectedTextCallback` asynchronous callback function.</span></span>
    
- <span data-ttu-id="a956f-223">Когда выполняется асинхронная функция обратного вызова, которая передается в качестве аргумента обратного вызова метода, она получает значение выбранного текста при возвращении обратного `selectTextCallback`  `getSelectedDataAsync` вызова.</span><span class="sxs-lookup"><span data-stu-id="a956f-223">When the  `selectTextCallback` asynchronous callback function that is passed as the _callback_ argument of the `getSelectedDataAsync` method executes, it gets the value of the selected text when the callback returns.</span></span> <span data-ttu-id="a956f-224">Оно получает это значение из выбранного аргумента _CallbackText_ (который имеет тип [AsyncResult)](/javascript/api/office/office.asyncresult)с помощью свойства значения возвращаемого [](/javascript/api/office/office.asyncresult#status) `AsyncResult` объекта.</span><span class="sxs-lookup"><span data-stu-id="a956f-224">It gets that value from the callback's _selectedText_ argument (which is of type [AsyncResult](/javascript/api/office/office.asyncresult)) by using the [value](/javascript/api/office/office.asyncresult#status) property of the returned `AsyncResult` object.</span></span>
    
- <span data-ttu-id="a956f-p120">Остальной код функции `selectedTextCallback` отправляет XML-веб-службе запрос на определения. Кроме того, он вызывает API-интерфейсы Microsoft Translator для получения URL-адреса WAV-файла с произношением выделенного слова.</span><span class="sxs-lookup"><span data-stu-id="a956f-p120">The rest of the code in the  `selectedTextCallback` function queries the XML web service for definitions. It also calls into the Microsoft Translator APIs to provide the URL of a .wav file that has the selected word's pronunciation.</span></span>
    
- <span data-ttu-id="a956f-227">Остальной код в файле Dictionary.js служит для отображения списка определений и ссылок на произношения в пользовательском интерфейсе HTML надстройки.</span><span class="sxs-lookup"><span data-stu-id="a956f-227">The remaining code in Dictionary.js displays the list of definitions and the pronunciation link in the add-in's HTML UI.</span></span>
    



```js
// The document the dictionary add-in is interacting with.
var _doc; 
// The last looked-up word, which is also the currently displayed word.
var lastLookup; 
// For demo purposes only!! Get an AppID if you intend to use the Pronunciation service for your feature.
var appID="3D8D4E1888B88B975484F0CA25CDD24AAC457ED8"; 

// The base URL for the OfficeDefinitions-conforming XML web service to query for definitions.
var xmlServiceUrl = "WebService.asmx/Define?Word="; 

// Initialize the add-in. 
// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    // Store a reference to the current document.
    _doc = Office.context.document; 
    // Check whether text is already selected.
    tryUpdatingSelectedWord(); 
    _doc.addHandlerAsync("documentSelectionChanged", tryUpdatingSelectedWord); //Add a handler to refresh when the user changes selection.
    });
}

// Executes when event is raised on user's selection changes, and at initialization time. 
// Gets the current selection and passes that to asynchronous callback method.
function tryUpdatingSelectedWord() {
    _doc.getSelectedDataAsync(Office.CoercionType.Text, selectedTextCallback); 
}

// Async callback that executes when the add-in gets the user's selection.
// Determines whether anything should be done. If so, it makes requests that will be passed to various functions.
function selectedTextCallback(selectedText) {
    selectedText = $.trim(selectedText.value);
    // Be sure user has selected text. The SelectionChanged event is raised every time the user moves the cursor, even if no selection.
    if (selectedText != "") { 
        // Check whether user selected the same word the pane is currently displaying to avoid unnecessary web calls.
        if (selectedText != lastLookup) { 
            // Update the lastLookup variable.
            lastLookup = selectedText; 
            // Set the "headword" span to the word you looked up.
            $("#headword").text(selectedText); 
            // AJAX request to get definitions for the selected word; pass that to refreshDefinitions.
            $.ajax(xmlServiceUrl + selectedText, { dataType: 'xml', success: refreshDefinitions, error: errorHandler }); 
            // AJAX request to the Microsoft Translator APIs. Gets the URL of a WAV file with pronunciation, which is passed to refreshPronunciation. See http://www.microsofttranslator.com/dev for details.
            $.ajax("http://api.microsofttranslator.com/V2/Ajax.svc/Speak?oncomplete=refreshPronunciation&amp;appId=" + appID + "&amp;text=" + selectedText + "&amp;language=en-us", { dataType: 'script', success: null, error: errorHandler }); 
        }
    }
}

// This function is called when the add-in gets back the definitions target word.
// It removes the old definitions and replaces them with the definitions for the current word.
// It also sets the "See More" link.
function refreshDefinitions(data, textStatus, jqXHR) {
    $(".definition").remove();
    // Make a new list item for each returned definition that was returned, set the CSS class, and append it to the definitions div.
    $(data).find("Definition").each(function () {
        $(document.createElement("li")).text($(this).text()).addClass("definition").appendTo($("#definitions"));
    });
    $("#SeeMoreLink").attr("href", $(data).find("SeeMoreURL").text()); //Change the "See More" link to direct to the correct URL.
}

// This function is called when the add-in gets back the link to the pronunciation
// to set the "Pronounce" link to the URL of the .WAV file.
function refreshPronunciation(data) {
    $("#pronunciationLink").attr("href", data);
}

// Basic error handler that writes to a div with id='message'.
function errorHandler(jqXHR, textStatus, errorThrown) {
    document.getElementById('message').innerText += errorThrown;
}

```
