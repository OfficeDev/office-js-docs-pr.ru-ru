---
title: Проверка манифеста и устранение связанных с ним неполадок
description: Используйте эти методы для проверки манифеста надстройки Office.
ms.date: 12/04/2017
ms.openlocfilehash: 51d644f7cfb7fbad5c9b66be41dc57015202b9be
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/19/2018
ms.locfileid: "25639989"
---
# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="f243a-103">Проверка манифеста и устранение связанных с ним неполадок</span><span class="sxs-lookup"><span data-stu-id="f243a-103">Validate and troubleshoot issues with your manifest</span></span>

<span data-ttu-id="f243a-104">Проверить манифест надстройки Office и устранить связанные с ним неполадки можно с помощью указанных ниже методов.</span><span class="sxs-lookup"><span data-stu-id="f243a-104">Use these methods to validate and troubleshoot issues in your Office Add-ins manifest:</span></span> 

- [<span data-ttu-id="f243a-105">Проверка манифеста с помощью средства проверки надстроек Office</span><span class="sxs-lookup"><span data-stu-id="f243a-105">Validate your manifest with the Office Add-in Validator</span></span>](#validate-your-manifest-with-the-office-add-in-validator)   
- [<span data-ttu-id="f243a-106">Проверка манифеста на соответствие схеме XML</span><span class="sxs-lookup"><span data-stu-id="f243a-106">Validate your manifest against the XML schema</span></span>](#validate-your-manifest-against-the-xml-schema)
- [<span data-ttu-id="f243a-107">Отладка манифеста надстройки с помощью журнала среды выполнения</span><span class="sxs-lookup"><span data-stu-id="f243a-107">Use runtime logging to debug your add-in manifest</span></span>](#use-runtime-logging-to-debug-your-add-in-manifest)


## <a name="validate-your-manifest-with-the-office-add-in-validator"></a><span data-ttu-id="f243a-108">Проверка манифеста с помощью средства проверки надстроек Office</span><span class="sxs-lookup"><span data-stu-id="f243a-108">Validate your manifest with the Office Add-in Validator</span></span>

<span data-ttu-id="f243a-109">Чтобы убедиться, что файл манифеста правильно и полностью описывает надстройку Office, проверьте его с помощью [средства проверки надстроек Office](https://github.com/OfficeDev/office-addin-validator).</span><span class="sxs-lookup"><span data-stu-id="f243a-109">To help ensure that the manifest file that describes your Office Add-in is correct and complete, validate it against the [Office Add-in Validator](https://github.com/OfficeDev/office-addin-validator).</span></span>

### <a name="to-use-the-office-add-in-validator-to-validate-your-manifest"></a><span data-ttu-id="f243a-110">Как проверить манифест с помощью средства проверки надстроек Office</span><span class="sxs-lookup"><span data-stu-id="f243a-110">To use the Office Add-in Validator to validate your manifest</span></span>

1. <span data-ttu-id="f243a-111">Установите [Node.js](https://nodejs.org/download/).</span><span class="sxs-lookup"><span data-stu-id="f243a-111">Install [Node.js](https://nodejs.org/download/).</span></span> 

2. <span data-ttu-id="f243a-112">Откройте командную строку или терминал от имени администратора и глобально установите средство проверки надстроек, используя следующую команду:</span><span class="sxs-lookup"><span data-stu-id="f243a-112">Open a command prompt / terminal as an administrator, and install the Office Add-in Validator and its dependencies globally by using the following command:</span></span>

    ```bash
    npm install -g office-addin-validator
    ```
    
    > [!NOTE]
    > <span data-ttu-id="f243a-113">Если у вас уже установлено приложение Yo Office, обновите его до последней версии, при этом средство проверки будет установлено в виде зависимости.</span><span class="sxs-lookup"><span data-stu-id="f243a-113">If you already have Yo Office installed, upgrade to the latest version, and the validator will be installed as a dependency.</span></span>

3. <span data-ttu-id="f243a-p101">Выполните приведенную ниже команду для проверки манифеста. Вместо файла MANIFEST.XML укажите путь к XML-файлу манифеста.</span><span class="sxs-lookup"><span data-stu-id="f243a-p101">Run the following command to validate your manifest. Replace MANIFEST.XML with the path to the manifest XML file.</span></span>

    ```bash
    validate-office-addin MANIFEST.XML
    ```

## <a name="validate-your-manifest-against-the-xml-schema"></a><span data-ttu-id="f243a-116">Проверка манифеста на соответствие схеме XML</span><span class="sxs-lookup"><span data-stu-id="f243a-116">Validate your manifest against the XML schema</span></span>

<span data-ttu-id="f243a-p102">Это позволяет обеспечить соответствие файла манифеста правильной схеме (в том числе корректность пространства имен для используемых элементов). Если вы скопировали элементы из других примеров манифеста, еще раз проверьте, указаны ли **подходящие пространства имен.**. Вы можете проверить манифест, используя файлы [XML  Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas). Для этой проверки можно использовать средство проверки на соответствие схеме XML.</span><span class="sxs-lookup"><span data-stu-id="f243a-p102">To help ensure that the manifest file follows the correct schema, including any namespaces for the elements you are using. If you copied elements from other sample manifests double check you also **include the appropiate namespaces**. You can validate a manifest against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) files. You can use an XML schema validation tool to perform this validation.</span></span> 



### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a><span data-ttu-id="f243a-121">Как проверить манифест на соответствие схеме XML с помощью программы командной строки</span><span class="sxs-lookup"><span data-stu-id="f243a-121">To use a command-line XML schema validation tool to validate your manifest</span></span>

1.  <span data-ttu-id="f243a-122">Установите [tar](https://www.gnu.org/software/tar/) и [libxml](http://xmlsoft.org/FAQ.html), если вы еще этого не сделали.</span><span class="sxs-lookup"><span data-stu-id="f243a-122">Install [tar](https://www.gnu.org/software/tar/) and [libxml](http://xmlsoft.org/FAQ.html), if you haven't already.</span></span>

2.  <span data-ttu-id="f243a-p103">Выполните указанную ниже команду. Вместо `XSD_FILE` укажите путь к XSD-файлу манифеста, а вместо `XML_FILE` — путь к XML-файлу манифеста.</span><span class="sxs-lookup"><span data-stu-id="f243a-p103">Run the following command. Replace `XSD_FILE` with the path to the manifest XSD file, and replace `XML_FILE` with the path to the manifest XML file.</span></span>
    
    ```bash
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="use-runtime-logging-to-debug-your-add-in"></a><span data-ttu-id="f243a-125">Используйте журнал среды выполнения для отладки вашей надстройки</span><span class="sxs-lookup"><span data-stu-id="f243a-125">Use runtime logging to debug your add-in manifest</span></span> 

<span data-ttu-id="f243a-p104">Можно использовать журнал среды выполнения для отладки надстройки в манифесте, а также нескольких ошибок установки. Эта функция может помочь вам определять и устранять проблемы с манифестом, которые не обнаруживаются при проверке схемы XSD, например несоответствие идентификаторов ресурсов. Ведение журнала в среде выполнения особенно полезно для отладки надстроек, которые реализуют команды надстройки и пользовательские функции Excel.</span><span class="sxs-lookup"><span data-stu-id="f243a-p104">You can use runtime logging to debug your add-in's manifest as well as several installation errors. This feature can help you identify and fix issues with your manifest that are not detected by XSD schema validation, such as a mismatch between resource IDs. Runtime logging is particularly  useful for debugging add-ins that implement add-in commands and Excel custom functions.</span></span>   

> [!NOTE]
> <span data-ttu-id="f243a-129">В настоящее время функция ведения журнала в среде выполнения доступна для классических приложений Office 2016.</span><span class="sxs-lookup"><span data-stu-id="f243a-129">The runtime logging feature is currently available for Office 2016 desktop.</span></span>

### <a name="to-turn-on-runtime-logging"></a><span data-ttu-id="f243a-130">Как включить ведение журнала в среде выполнения</span><span class="sxs-lookup"><span data-stu-id="f243a-130">To turn on runtime logging</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f243a-p105">Ведение журнала в среде выполнения снижает производительность. Включайте его, только когда нужно исправить ошибки в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="f243a-p105">Runtime Logging affects performance. Turn it on only when you need to debug issues with your add-in manifest.</span></span>

<span data-ttu-id="f243a-133">Чтобы включить ведение журнала в среде выполнения:</span><span class="sxs-lookup"><span data-stu-id="f243a-133">To turn on runtime logging:</span></span>

1. <span data-ttu-id="f243a-134">Убедитесь, что у вас установлена сборка Office 2016 **16.0.7019** или выше.</span><span class="sxs-lookup"><span data-stu-id="f243a-134">Make sure that you are running Office 2016 desktop build **16.0.7019** or later.</span></span> 

2. <span data-ttu-id="f243a-135">Добавьте раздел реестра `RuntimeLogging` в раздел `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\`.</span><span class="sxs-lookup"><span data-stu-id="f243a-135">Add the `RuntimeLogging` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\`.</span></span> 

3. <span data-ttu-id="f243a-p106">В качестве значения по умолчанию задайте полный путь к файлу, в который будет записываться журнал. Пример приведен в архиве [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span><span class="sxs-lookup"><span data-stu-id="f243a-p106">Set the default value of the key to the full path of the file where you want the log to be written. For an example, see [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span></span> 

    > [!NOTE]
    > <span data-ttu-id="f243a-138">Необходим готовый каталог, в котором будет создан файл журнала, и соответствующее разрешение на запись.</span><span class="sxs-lookup"><span data-stu-id="f243a-138">The directory in which the log file will be written must already exist, and you must have write permissions to it.</span></span> 
 
<span data-ttu-id="f243a-p107">На рисунке ниже показано, как должен выглядеть реестр. Чтобы отключить функцию, удалите из реестра раздел `RuntimeLogging`.</span><span class="sxs-lookup"><span data-stu-id="f243a-p107">The following image shows what the registry should look like. To turn the feature off, remove the `RuntimeLogging` key from the registry.</span></span> 

![Снимок экрана "Редактор реестра" с разделом реестра RuntimeLogging](http://i.imgur.com/Sa9TyI6.png)


### <a name="to-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="f243a-142">Как устранить проблемы с манифестом</span><span class="sxs-lookup"><span data-stu-id="f243a-142">To troubleshoot issues with your manifest</span></span>

<span data-ttu-id="f243a-143">Чтобы устранить проблемы с загрузкой надстройки, используя журнал среды выполнения:</span><span class="sxs-lookup"><span data-stu-id="f243a-143">To use runtime logging to troubleshoot issues loading an add-in:</span></span>
 
1. <span data-ttu-id="f243a-144">[Загрузите неопубликованную надстройку](sideload-office-add-ins-for-testing.md) для тестирования.</span><span class="sxs-lookup"><span data-stu-id="f243a-144">[Sideload your add-in](sideload-office-add-ins-for-testing.md) for testing.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="f243a-145">Рекомендуем загружать только тестируемую надстройку, чтобы уменьшить количество сообщений в файле журнала.</span><span class="sxs-lookup"><span data-stu-id="f243a-145">We recommend that you sideload only the add-in that you are testing to minimize the number of messages in the log file.</span></span>

2. <span data-ttu-id="f243a-146">Если ничего не происходит и надстройка не отображается в диалоговом окне надстроек, откройте файл журнала.</span><span class="sxs-lookup"><span data-stu-id="f243a-146">If nothing happens and you don't see your add-in (and it's not appearing in the add-ins dialog box), open the log file.</span></span>

3. <span data-ttu-id="f243a-p108">Выполните в этом файле поиск по идентификатору надстройки, определенному в манифесте. В файле журнала этот идентификатор отмечен как `SolutionId`.</span><span class="sxs-lookup"><span data-stu-id="f243a-p108">Search the log file for your add-in ID, which you define in your manifest. In the log file, this ID is labeled `SolutionId`.</span></span> 

<span data-ttu-id="f243a-p109">В приведенном ниже примере файл журнала определяет элемент управления, указывающий на несуществующий файл ресурсов. В этом примере необходимо исправить опечатку в манифесте или добавить недостающий ресурс.</span><span class="sxs-lookup"><span data-stu-id="f243a-p109">In the following example, the log file identifies a control that points to a resource file that doesn't exist. For this example, the fix would be to correct the typo in the manifest or to add the missing resource.</span></span>

![Снимок экрана с файлом журнала, содержащим запись, которая указывает на несуществующий идентификатор ресурса.](http://i.imgur.com/f8bouLA.png) 

### <a name="known-issues-with-runtime-logging"></a><span data-ttu-id="f243a-152">Известные проблемы с ведением журнала в среде выполнения</span><span class="sxs-lookup"><span data-stu-id="f243a-152">Known issues with runtime logging</span></span>

<span data-ttu-id="f243a-p110">В файле журнала могут встречаться непонятные или неправильно классифицированные сообщения. Например:</span><span class="sxs-lookup"><span data-stu-id="f243a-p110">You might see messages in the log file that are confusing or that are classified incorrectly. For example:</span></span>

- <span data-ttu-id="f243a-155">сообщение `Medium Current host not in add-in's host list` с дополнением `Unexpected Parsed manifest targeting different host` неправильно классифицируется как ошибка.</span><span class="sxs-lookup"><span data-stu-id="f243a-155">The message `Medium Current host not in add-in's host list` followed by `Unexpected Parsed manifest targeting different host` is incorrectly classified as an error.</span></span>

- <span data-ttu-id="f243a-156">Если появится сообщение `Unexpected Add-in is missing required manifest fields DisplayName`, не содержащее SolutionId, то ошибка, скорее всего, не связана с надстройкой, отладка которой выполняется.</span><span class="sxs-lookup"><span data-stu-id="f243a-156">If you see the message `Unexpected Add-in is missing required manifest fields DisplayName` and it doesn't contain a SolutionId, the error is most likely not related to the add-in you are debugging.</span></span> 

- <span data-ttu-id="f243a-p111">Все сообщения `Monitorable` являются ожидаемыми ошибками с точки зрения системы. Иногда они указывают на проблему с манифестом, например опечатку в элементе, которая была пропущена, но не привела к сбою.</span><span class="sxs-lookup"><span data-stu-id="f243a-p111">Any `Monitorable` messages are expected errors from a system point of view. Sometimes they indicate an issue with your manifest, such as a misspelled element that was skipped but didn't cause the manifest to fail.</span></span> 

## <a name="clear-the-office-cache"></a><span data-ttu-id="f243a-159">Очистка кэша Office</span><span class="sxs-lookup"><span data-stu-id="f243a-159">Clear the Office cache</span></span>

<span data-ttu-id="f243a-160">Если внесенные в манифест изменения (например, имена значков кнопок на ленте или текст команд надстроек) не вступили в силу, попробуйте очистить кэш Office на своем компьютере.</span><span class="sxs-lookup"><span data-stu-id="f243a-160">If changes you've made in the manifest, such as file names of ribbon button icons, or text of add-in commands, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="f243a-161">Для Windows:</span><span class="sxs-lookup"><span data-stu-id="f243a-161">For Windows:</span></span>
<span data-ttu-id="f243a-162">Удалите содержимое папки `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="f243a-162">Delete the content of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="f243a-163">Для Mac</span><span class="sxs-lookup"><span data-stu-id="f243a-163">For Mac:</span></span>
<span data-ttu-id="f243a-164">Удалите содержимое папки `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="f243a-164">Delete the content of the folder `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>

#### <a name="for-ios"></a><span data-ttu-id="f243a-165">Для iOS</span><span class="sxs-lookup"><span data-stu-id="f243a-165">For iOS:</span></span>
<span data-ttu-id="f243a-p112">Для принудительной перезагрузки вызовите метод JavaScript `window.location.reload(true)` в надстройке. Вы также можете переустановить Office.</span><span class="sxs-lookup"><span data-stu-id="f243a-p112">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="f243a-168">См. также</span><span class="sxs-lookup"><span data-stu-id="f243a-168">See also</span></span>

- [<span data-ttu-id="f243a-169">XML-манифест надстройки Office</span><span class="sxs-lookup"><span data-stu-id="f243a-169">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="f243a-170">Загрузка неопубликованных надстроек Office для тестирования</span><span class="sxs-lookup"><span data-stu-id="f243a-170">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="f243a-171">Отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="f243a-171">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
