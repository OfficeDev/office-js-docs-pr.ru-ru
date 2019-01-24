---
title: Обновление библиотеки API JavaScript для Office до последней версии и схемы манифеста надстройки до версии 1.1
description: Обновление до версии 1.1 файлов JavaScript (Office.js и JS-файлов приложения) и файла проверки манифеста надстройки в проекте надстройки Office.
ms.date: 12/12/2018
localization_priority: Normal
ms.openlocfilehash: 20c6c6362aa09926e967e52edfe6be69a09edb18
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/23/2019
ms.locfileid: "29387725"
---
# <a name="update-to-the-latest-javascript-api-for-office-library-and-version-11-add-in-manifest-schema"></a>Обновление библиотеки API JavaScript для Office до последней версии и схемы манифеста надстройки до версии 1.1

В этой статье рассказывается, как обновить файлы JavaScript (Office.js и JS-файлы для конкретной надстройки) и файл проверки манифеста надстройки в проекте надстройки Office до версии 1.1.

> [!NOTE]
> Проекты, создаваемые в Visual Studio 2017, уже используют версию 1.1. Однако для версии 1.1 периодически выпускаются незначительные обновления, которые можно применить с помощью методов, описанных в этой статье.

## <a name="use-the-most-up-to-date-project-files"></a>Использование последних версий файлов в проекте

Если для разработки надстройки вы используете Visual Studio, то чтобы можно было применять [самые новые элементы API](https://docs.microsoft.com/office/dev/add-ins/reference/what's-changed-in-the-javascript-api-for-office) в API JavaScript для Office и [возможности манифеста надстройки версии 1.1](../develop/add-in-manifests.md) (который проверяется на соответствие offappmanifest-1.1.xsd), вам потребуется скачать Visual Studio 2017. Чтобы скачать Visual Studio 2017, перейдите на [страницу интегрированной среды разработки Visual Studio](https://visualstudio.microsoft.com/vs/). Во время установки потребуется выбрать рабочую нагрузку разработки Office и SharePoint.

Если вы используете текстовый редактор или другую интегрированную среду разработки, отличную от Visual Studio, чтобы разработать надстройка, обновите ссылки на CDN для файла Office.js и версию схемы, на которую ссылается манифест приложения для Office.

Чтобы запустить надстройку, разработанную с использованием новых и обновленных компонентов манифеста надстройки и интерфейса API Office.js, ваши клиенты должны использовать локальные продукты Office 2013 с пакетом обновления 1 (SP1) или более поздней версии, а также при необходимости SharePoint Server 2013 с пакетом обновления 1 (SP1) и связанными серверными продуктами, Пакет обновления 1 (SP1) для Exchange Server 2013 или аналогичные размещенные в сети продукты: Office 365, SharePoint Online и Exchange Online.

Сведения о том, как скачать Office, SharePoint и Exchange с пакетом обновления 1, см. в следующих статьях:

- [Список всех пакетов обновления 1 (SP1) для Microsoft Office 2013 и связанных продуктов для настольных систем](https://support.microsoft.com/kb/2850036)
    
- [Список всех пакетов обновления 1 (SP1) для Microsoft SharePoint Server 2013 и связанных серверных продуктов](https://support.microsoft.com/kb/2850035)
    
- [Описание пакета обновления 1 для Exchange Server 2013](https://support.microsoft.com/kb/2926248)
    

## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a>Обновление проекта надстройки Office, созданного в Visual Studio

Для проектов, созданных до выпуска версии 1.1 библиотеки JavaScript API для Office и схемы манифеста надстройки, вы можете обновить файлы проекта, используя **диспетчер пакетов NuGet**, а затем добавить ссылки на них в HTML-страницы надстройки. 

Обратите внимание, что процесс обновления применяется к _проектам по отдельности_. Вам потребуется повторить его для каждого проекта надстройки, в котором вы хотите использовать Office.js и схему манифеста надстройки версии 1.1.


### <a name="update-the-javascript-api-for-office-library-files-in-your-project-to-the-newest-release"></a>Обновление файлов библиотеки API JavaScript для Office в проекте до последней версии
Обновить файлы библиотеки Office до последней версии можно с помощью указанных ниже действий. В них используется Visual Studio 2017, но они аналогичны для Visual Studio 2015.

1. В Visual Studio 2017 откройте или создайте проект **Надстройка Office**.    
2. Выберите **Средства** > **Диспетчер пакетов NuGet** > **Управление пакетами Nuget для решения**.
3. В **диспетчере пакетов NuGet** выберите **nuget.org** для параметра **Источник пакетов**.
4. Выберите вкладку **Обновления**.
5. Выберите Microsoft.Office.js.
6. В области слева выберите **Обновить** и завершите обновление пакета.

Вам потребуется выполнить несколько дополнительных действий, чтобы завершить обновление. В теге **head** HTML-страниц надстройки закомментируйте или удалите все ссылки на скрипт office.js и добавьте ссылки на обновленную библиотеку API JavaScript для Office, как показано ниже.
    
  ```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
  ```

   > [!NOTE] 
   > `/1/` в `office.js` в URL-адресе CDN указывает на то, что необходимо использовать последний добавочный выпуск Office.js версии 1.


### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>Обновление схемы манифеста в проекте до версии 1.1

В файле манифеста надстройки обновите атрибут **xmlns** элемента **OfficeApp**, заменив значение версии на `1.1` и оставив все атрибуты, кроме **xmlns**, без изменений.
    
```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp" 
    xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE] 
> После обновления схемы манифеста надстройки до версии 1.1 вам потребуется удалить элементы   **Capabilities** и **Capability** и заменить их либо элементами [Hosts](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/hosts) и [Host](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/host), либо [элементами Requirements и Requirement](specify-office-hosts-and-api-requirements.md).

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a>Обновление проекта надстройки Office, созданного с помощью текстового редактора или другой среды IDE

Если вы создали проект до выпуска схемы манифеста надстройки и API JavaScript для Office версии 1.1, обновите HTML-страницы вашей надстройки, чтобы они ссылались на CDN библиотеки версии 1.1, а также обновите файл манифеста надстройки, чтобы использовалась схема версии 1.1. 

Процесс обновления применяется к _проектам по отдельности_. Вам потребуется повторить его для каждого проекта надстройки, в котором вы хотите использовать файл Office.js и схему манифеста надстройки версии 1.1.

Вам не нужны локальные копии файлов API JavaScript для Office (Office.js и JS-файлов для конкретной надстройки), чтобы разрабатывать надстройку Office (ссылки на CDN для Office.js позволяют скачивать необходимые файлы во время выполнения). Если вам нужны файлы библиотеки, то вы можете скачать их с помощью [служебной программы командной строки NuGet](https://docs.nuget.org/consume/installing-nuget) и `Install-Package Microsoft.Office.js`.

> [!NOTE] 
> Чтобы получить копию XSD (определения схемы XML) для манифеста надстройки версии 1.1, см. статью [Справочник по схеме манифестов надстроек Office (версия 1.1)](../develop/add-in-manifests.md).


### <a name="update-the-javascript-api-for-office-library-files-in-your-project-to-use-the-newest-release"></a>Обновление файлов библиотеки API JavaScript для Office в проекте до последней версии

1. Откройте HTML-страницы надстройки в текстовом редакторе или интегрированной среде разработки.
    
2. В теге **head** HTML-страниц надстройки закомментируйте или удалите все ссылки на скрипт office.js и добавьте ссылки на обновленную библиотеку API JavaScript для Office, как показано ниже.
    
    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    ```

   > [!NOTE] 
   > `/1/` перед `office.js` в URL-адресе CDN указывает на то, что необходимо использовать последний добавочный выпуск Office.js версии 1.   

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>Обновление схемы манифеста в проекте до версии 1.1

В файле манифеста надстройки обновите атрибут **xmlns** элемента **OfficeApp**, заменив значение версии на `1.1` и оставив все атрибуты, кроме **xmlns**, без изменений.
    
```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp" 
    xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE] 
> После обновления схемы манифеста надстройки до версии 1.1 вам потребуется удалить элементы   **Capabilities** и **Capability** и заменить их либо элементами [Hosts](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/hosts) и [Host](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/host), либо [элементами Requirements и Requirement](specify-office-hosts-and-api-requirements.md).
    

## <a name="see-also"></a>См. также

- [Указание ведущих приложений Office и элементов API](specify-office-hosts-and-api-requirements.md) 
- [Общие сведения об интерфейсе API JavaScript для Office](understanding-the-javascript-api-for-office.md)    
- [API JavaScript для Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)   
- [Справка по схеме для манифестов надстроек Office (версия 1.1)](../develop/add-in-manifests.md)
    
