---
title: Обновление до последней библиотеки API JavaScript для Office и схемы манифеста надстройки версии 1.1
description: Обновление до версии 1.1 файлов JavaScript (Office.js и JS-файлов приложения) и файла проверки манифеста надстройки в проекте надстройки Office.
ms.date: 01/14/2022
ms.localizationpriority: medium
ms.openlocfilehash: 32fcadb6a36ca540a799f8d6a5dfa671ee5e5de8
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660202"
---
# <a name="update-to-the-latest-office-javascript-api-library-and-version-11-add-in-manifest-schema"></a>Обновление до последней библиотеки API JavaScript для Office и схемы манифеста надстройки версии 1.1

В этой статье рассказывается, как обновить файлы JavaScript (Office.js и JS-файлы для конкретной надстройки) и файл проверки манифеста надстройки в проекте надстройки Office до версии 1.1.

> [!NOTE]
> Проекты, созданные в Visual Studio 2019, уже будут использовать версию 1.1. Однако для версии 1.1 периодически выпускаются незначительные обновления, которые можно применить с помощью методов, описанных в этой статье.

## <a name="use-the-most-up-to-date-project-files"></a>Использование последних версий файлов в проекте

Если для разработки надстройки используется Visual Studio, для использования новых элементов API API API JavaScript для Office и функций манифеста надстройки версии [1.1](../develop/add-in-manifests.md) (который проверяется на соответствие offappmanifest-1.1.xsd), необходимо скачать Visual Studio 2019. Чтобы скачать Visual Studio 2019, перейдите на страницу [интегрированной среды разработки Visual Studio](https://visualstudio.microsoft.com/vs/). Во время установки потребуется выбрать рабочую нагрузку разработки Office и SharePoint.

Если для разработки надстройки используется текстовый редактор или интегрированная среда разработки, отличные от Visual Studio, необходимо обновить ссылки на сеть доставки содержимого (CDN) для Office.js и версию схемы, на которую ссылается манифест надстройки.

Чтобы запустить надстройку, разработанную с помощью нового и обновленного API Office.js и функций манифеста надстройки, ваши клиенты должны запускать локальные продукты Office 2013 с пакетом обновления 1 (SP1) или более поздней версии, а также, если это применимо, SharePoint Server 2013 с пакетом обновления 1 (SP1) и связанные серверные продукты, Exchange Server 2013 с пакетом обновления 1 (SP1) или эквивалентные веб-продукты, размещенные в Интернете: Microsoft 365, SharePoint Online и Exchange Online.

Сведения о том, как скачать Office, SharePoint и Exchange с пакетом обновления 1, см. в следующих статьях:

- [Список всех пакетов обновления 1 (SP1) для Microsoft Office 2013 и связанных продуктов для настольных систем](https://support.microsoft.com/kb/2850036)

- [Список всех пакетов обновления 1 (SP1) для Microsoft SharePoint Server 2013 и связанных серверных продуктов](https://support.microsoft.com/kb/2850035)

- [Описание пакета обновления 1 для Exchange Server 2013](https://support.microsoft.com/kb/2926248)

## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a>Обновление проекта надстройки Office, созданного в Visual Studio

Для проектов, созданных до выпуска версии 1.1 API JavaScript для Office и схемы манифеста надстройки, можно обновить файлы проекта с помощью диспетчера пакетов **NuGet**, а затем обновить HTML-страницы надстройки, чтобы ссылаться на них.

Обратите внимание, что процесс обновления применяется к _проектам по отдельности_. Вам потребуется повторить его для каждого проекта надстройки, в котором вы хотите использовать Office.js и схему манифеста надстройки версии 1.1.

### <a name="update-the-office-javascript-api-library-files-in-your-project-to-the-newest-release"></a>Обновление файлов библиотеки API JavaScript для Office в проекте до самого нового выпуска

Ниже описано, как обновить файлы Office.js до последней версии. В этих шагах используется Visual Studio 2019, но они аналогичны предыдущим версиям Visual Studio.

1. В Visual Studio 2019 откройте или создайте проект **надстройки Office** .
2. Выберите **"Инструменты** > **" диспетчера пакетов** > **NuGet для управления пакетами NuGet для решения**.
3. Выберите вкладку **Обновления**.
4. Выберите Microsoft.Office.js. Убедитесь, что источник пакета находится **из nuget.org**.
5. В левой области выберите " **Установить"** и завершите процесс обновления пакета.

Вам потребуется выполнить несколько дополнительных действий, чтобы завершить обновление. В **теге** заголовка HTML-страниц надстройки закомментировать или удалить все существующие ссылки на скрипты office.js и ссылаться на обновленную библиотеку API JavaScript для Office следующим образом:

  ```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
  ```

   > [!NOTE]
   > `/1/` в `office.js` в URL-адресе CDN указывает на то, что необходимо использовать последний добавочный выпуск Office.js версии 1.

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>Обновление схемы манифеста в проекте до версии 1.1

В файле манифеста надстройки обновите атрибут **xmlns** **\<OfficeApp\>** `1.1` элемента, изменив значение версии на (оставив атрибуты, отличные от **атрибута xmlns** , без изменений).

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> После обновления версии схемы манифеста надстройки до версии 1.1 необходимо удалить элементы **Capabilities** и **Capability** и заменить их элементами [Hosts](/javascript/api/manifest/hosts) и [Host](/javascript/api/manifest/host) или элементами [Requirements и Requirement](specify-office-hosts-and-api-requirements.md).

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a>Обновление проекта надстройки Office, созданного с помощью текстового редактора или другой среды IDE

Для проектов, созданных до выпуска версии 1.1 API JavaScript для Office и схемы манифеста надстройки, необходимо обновить HTML-страницы надстройки для ссылки на CDN библиотеки версии 1.1 и обновить файл манифеста надстройки для использования схемы версии 1.1.

Процесс обновления применяется к _проектам по отдельности_. Вам потребуется повторить его для каждого проекта надстройки, в котором вы хотите использовать файл Office.js и схему манифеста надстройки версии 1.1.

Вам не нужны локальные копии файлов API JavaScript для Office (Office.js и файлы .js для приложений) для разработки надстройкиOffice (ссылка на CDN для Office.js загружает необходимые файлы во время выполнения), но если требуется локальная копия файлов библиотеки, можно использовать служебную программу [NuGet Command-Line](https://docs.nuget.org/consume/installing-nuget) `Install-Package Microsoft.Office.js` и команду для их скачивания.

> [!NOTE]
> Чтобы получить копию XSD (определения схемы XML) для манифеста надстройки версии 1.1, см. статью [Справочник по схеме манифестов надстроек Office (версия 1.1)](../develop/add-in-manifests.md).

### <a name="update-the-office-javascript-api-library-files-in-your-project-to-use-the-newest-release"></a>Обновление файлов библиотеки API JavaScript для Office в проекте для использования самого нового выпуска

1. Откройте HTML-страницы надстройки в текстовом редакторе или интегрированной среде разработки.

2. В **теге** заголовка HTML-страниц надстройки закомментировать или удалить все существующие ссылки на скрипты office.js и ссылаться на обновленную библиотеку API JavaScript для Office следующим образом:

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
    ```

   > [!NOTE]
   > `/1/` перед `office.js` в URL-адресе CDN указывает на то, что необходимо использовать последний добавочный выпуск Office.js версии 1.

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>Обновление схемы манифеста в проекте до версии 1.1

В файле манифеста надстройки обновите атрибут **xmlns** **\<OfficeApp\>** `1.1` элемента, изменив значение версии на (оставив атрибуты, отличные от **атрибута xmlns** , без изменений).

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> После обновления версии схемы манифеста надстройки до версии 1.1 необходимо удалить элементы **Capabilities** и **Capability** и заменить их элементами [Hosts](/javascript/api/manifest/hosts) и [Host](/javascript/api/manifest/host) или элементами [Requirements и Requirement](specify-office-hosts-and-api-requirements.md).

## <a name="see-also"></a>См. также

- [Указание приложений Office и требований К API](specify-office-hosts-and-api-requirements.md) ]
- [Общие сведения об API JavaScript для Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript для Office](../reference/javascript-api-for-office.md)
- [Справка по схеме для манифестов надстроек Office (версия 1.1)](../develop/add-in-manifests.md)
