---
title: Обновление библиотеки API JavaScript для Office до последней версии и схемы манифеста надстройки до версии 1.1
description: ''
ms.date: 12/04/2017
---

# <a name="update-to-the-latest-javascript-api-for-office-library-and-version-11-add-in-manifest-schema"></a>Обновление библиотеки API JavaScript для Office до последней версии и схемы манифеста надстройки до версии 1.1

В этой статье рассказывается, как обновить файлы JavaScript (Office.js и JS-файлы для конкретной надстройки) и файл проверки манифеста надстройки в проекте надстройки Office до версии 1.1.

## <a name="use-the-most-up-to-date-project-files"></a>Использование последних версий файлов в проекте

Если для разработки надстройки вы используете Visual Studio, то чтобы можно было применять [самые новые элементы API](https://dev.office.com/reference/add-ins/what's-changed-in-the-javascript-api-for-office) в API JavaScript для Office и [возможности манифеста надстройки версии 1.1](../develop/add-in-manifests.md) (который проверяется на соответствие offappmanifest-1.1.xsd), вам потребуется скачать и установить [Visual Studio 2015 и последнюю версию Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs).

Если вы используете текстовый редактор или другую интегрированную среду разработки, отличную от Visual Studio, чтобы разработать надстройка, обновите ссылки на CDN для файла Office.js и версию схемы, на которую ссылается манифест приложения для Office.

Чтобы запустить надстройку, разработанную с использованием новых и обновленных компонентов манифеста надстройки и интерфейса API Office.js, ваши клиенты должны использовать локальные продукты Office 2013 с пакетом обновления 1 (SP1) или более поздней версии, а также при необходимости SharePoint Server 2013 с пакетом обновления 1 (SP1) и связанными серверными продуктами, Пакет обновления 1 (SP1) для Exchange Server 2013 или аналогичные размещенные в сети продукты: Office 365, SharePoint Online и Exchange Online.

Сведения о том, как скачать Office, SharePoint и Exchange с пакетом обновления 1, см. в следующих статьях:

- [Список всех пакетов обновления 1 (SP1) для Microsoft Office 2013 и связанных продуктов для настольных систем](http://support.microsoft.com/kb/2850036)
    
- [Список всех пакетов обновления 1 (SP1) для Microsoft SharePoint Server 2013 и связанных серверных продуктов](http://support.microsoft.com/kb/2850035)
    
- [Описание пакета обновления 1 для Exchange Server 2013](http://support.microsoft.com/kb/2926248)
    

## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a>Обновление проекта надстройки Office, созданного в Visual Studio

Для проектов, созданных до выпуска версии 1.1 библиотеки JavaScript API для Office и схемы манифеста надстройки, вы можете обновить файлы проекта, используя **диспетчер пакетов NuGet**, а затем добавить ссылки на них в HTML-страницы надстройки. 

Обратите внимание, что процесс обновления применяется к _проектам по отдельности_. Вам потребуется повторить его для каждого проекта надстройки, в котором вы хотите использовать Office.js и схему манифеста надстройки версии 1.1.


### <a name="update-the-javascript-api-for-office-library-files-in-your-project-to-the-newest-release"></a>Обновление файлов библиотеки API JavaScript для Office в проекте до последней версии


1. В Visual Studio 2015 откройте или создайте проект **Надстройка Office**.
    
      - В расположенной слева области щелкните **Обновить** и завершите процесс обновления пакета.
    
      - Перейдите к этапу 6.
    
2. Выберите **Средства**  >  **Диспетчер пакетов NuGet**  >  **Управление пакетами Nuget для решения**.
    
3. В **диспетчере пакетов NuGet** выберите **nuget.org** в качестве **источника пакетов** и **Доступны обновления** в поле **Фильтр**. Затем выберите файл Microsoft.Office.js.
    
4. В области слева выберите **Обновить** и завершите обновление пакета.
    
5. В теге **head** HTML-страниц надстройки закомментируйте или удалите все ссылки на скрипт office.js и добавьте ссылки на обновленную библиотеку API JavaScript для Office, как показано ниже.
    
    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    ```

    > **ПРИМЕЧАНИЕ.** Цифра `/1/` перед `office.js` в URL-адресе CDN указывает на то, что необходимо использовать последний накопительный выпуск Office.js версии 1.   


### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>Обновление схемы манифеста в проекте до версии 1.1

В файле манифеста надстройки обновите атрибут **xmlns** элемента **OfficeApp**, заменив значение версии на `1.1` и оставив все атрибуты, кроме **xmlns**, без изменений:
    
```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp" 
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> **ПРИМЕЧАНИЕ.** После обновления схемы манифеста надстройки до версии 1.1 вам потребуется удалить элементы **Capabilities** и **Capability** и заменить их элементами [Hosts](https://dev.office.com/reference/add-ins/manifest/hosts) и [Host](https://dev.office.com/reference/add-ins/manifest/hosts) либо [элементами Requirements и Requirement](specify-office-hosts-and-api-requirements.md).

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a>Обновление проекта надстройки Office, созданного с помощью текстового редактора или другой среды IDE

Если вы создали проект до выпуска схемы манифеста надстройки и API JavaScript для Office версии 1.1, обновите HTML-страницы вашей надстройки, чтобы они ссылались на CDN библиотеки версии 1.1, а также обновите файл манифеста надстройки, чтобы использовалась схема версии 1.1. 

Процесс обновления применяется к _проектам по отдельности_. Вам потребуется повторить его для каждого проекта надстройки, в котором вы хотите использовать файл Office.js и схему манифеста надстройки версии 1.1.

Вам не нужны локальные копии файлов API JavaScript для Office (Office.js и JS-файлов для конкретной надстройки), чтобы разрабатывать надстройку Office (ссылки на CDN для Office.js позволяют скачивать необходимые файлы во время выполнения). Если вам нужны файлы библиотеки, то вы можете скачать их с помощью [служебной программы командной строки NuGet](http://docs.nuget.org/consume/installing-nuget) и `Install-Package Microsoft.Office.js`.

> **ПРИМЕЧАНИЕ.** Чтобы получить копию файла XSD для манифеста надстройки версии 1.1, см. запись в [типовых схемах для манифестов надстроек Office (версия 1.1)](../develop/add-in-manifests.md).


### <a name="update-the-javascript-api-for-office-library-files-in-your-project-to-use-the-newest-release"></a>Обновление файлов библиотеки API JavaScript для Office в проекте до последней версии

1. Откройте HTML-страницы надстройки в текстовом редакторе или интегрированной среде разработки.
    
2. В теге **head** HTML-страниц надстройки закомментируйте или удалите все ссылки на скрипт office.js и добавьте ссылки на обновленную библиотеку API JavaScript для Office, как показано ниже.
    
    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    ```

    > **ПРИМЕЧАНИЕ.** Цифра `/1/` перед `office.js` в URL-адресе CDN указывает на то, что необходимо использовать последний накопительный выпуск Office.js версии 1.   

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>Обновление схемы манифеста в проекте до версии 1.1

В файле манифеста надстройки обновите атрибут **xmlns** элемента **OfficeApp**, заменив значение версии на `1.1` и оставив все атрибуты, кроме **xmlns**, без изменений:
    
```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp" 
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> **ПРИМЕЧАНИЕ.** После обновления схемы манифеста надстройки до версии 1.1 вам потребуется удалить элементы **Capabilities** и **Capability** и заменить их элементами [Hosts](https://dev.office.com/reference/add-ins/manifest/hosts) и [Host](https://dev.office.com/reference/add-ins/manifest/hosts) либо [элементами Requirements и Requirement](specify-office-hosts-and-api-requirements.md).
    

## <a name="see-also"></a>См. также

- [Указание ведущих приложений Office и элементов API](specify-office-hosts-and-api-requirements.md) 
- [Общие сведения об интерфейсе API JavaScript для Office](understanding-the-javascript-api-for-office.md)    
- [API JavaScript для Office](https://dev.office.com/reference/add-ins/javascript-api-for-office)   
- [Справка по схеме для манифестов надстроек Office (версия 1.1)](../develop/add-in-manifests.md)
    
