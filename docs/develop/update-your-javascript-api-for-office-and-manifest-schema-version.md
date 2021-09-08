---
title: Обновление до последней Office библиотеки API JavaScript и схемы манифеста надстройки версии 1.1
description: Обновление до версии 1.1 файлов JavaScript (Office.js и JS-файлов приложения) и файла проверки манифеста надстройки в проекте надстройки Office.
ms.date: 10/11/2019
localization_priority: Normal
ms.openlocfilehash: b0536b4b55accd99e002e26c467572330ba72ae2
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58939092"
---
# <a name="update-to-the-latest-office-javascript-api-library-and-version-11-add-in-manifest-schema"></a>Обновление до последней Office библиотеки API JavaScript и схемы манифеста надстройки версии 1.1

В этой статье рассказывается, как обновить файлы JavaScript (Office.js и JS-файлы для конкретной надстройки) и файл проверки манифеста надстройки в проекте надстройки Office до версии 1.1.

> [!NOTE]
> Проекты, созданные Visual Studio 2019 г., уже будут использовать версию 1.1. Однако для версии 1.1 периодически выпускаются незначительные обновления, которые можно применить с помощью методов, описанных в этой статье.

## <a name="use-the-most-up-to-date-project-files"></a>Использование последних версий файлов в проекте

Если вы используете Visual Studio для разработки надстройки, для использования самых новых членов API API Office JavaScript и компонентов [v1.1](../develop/add-in-manifests.md) манифеста надстройки (которая проверяется в отношении offappmanifest-1.1.xsd), необходимо скачать Visual Studio 2019. Чтобы скачать Visual Studio 2019 г., см. страницу [Visual Studio IDE.](https://visualstudio.microsoft.com/vs/) Во время установки потребуется выбрать рабочую нагрузку разработки Office и SharePoint.

Если вы используете текстовый редактор или другую интегрированную среду разработки, отличную от Visual Studio, чтобы разработать надстройка, обновите ссылки на CDN для файла Office.js и версию схемы, на которую ссылается манифест приложения для Office.

Чтобы запустить надстройку, разработанную с использованием новых и обновленных Office.js API и функций манифеста надстройки, ваши клиенты должны запускать продукты sp1 2013 Office 2013 или более поздних версий локальной версии, а также, если это применимо, продукты SP1 SharePoint Server 2013 и связанные с ними серверы, Exchange Server 2013 Пакет обновления 1 (SP1) или эквивалентные веб-хост-продукты: Microsoft 365, SharePoint Online и Exchange Online.

Сведения о том, как скачать Office, SharePoint и Exchange с пакетом обновления 1, см. в следующих статьях:

- [Список всех пакетов обновления 1 (SP1) для Microsoft Office 2013 и связанных продуктов для настольных систем](https://support.microsoft.com/kb/2850036)

- [Список всех пакетов обновления 1 (SP1) для Microsoft SharePoint Server 2013 и связанных серверных продуктов](https://support.microsoft.com/kb/2850035)

- [Описание пакета обновления 1 для Exchange Server 2013](https://support.microsoft.com/kb/2926248)


## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a>Обновление проекта надстройки Office, созданного в Visual Studio

Для проектов, созданных перед выпуском v1.1 API Office JavaScript и схемы манифеста надстройки, можно обновить файлы проекта с помощью **NuGet диспетчер пакетов,** а затем обновить HTML-страницы надстройки для ссылки на них. 

Обратите внимание, что процесс обновления применяется к _проектам по отдельности_. Вам потребуется повторить его для каждого проекта надстройки, в котором вы хотите использовать Office.js и схему манифеста надстройки версии 1.1.

### <a name="update-the-office-javascript-api-library-files-in-your-project-to-the-newest-release"></a>Обновление Office файлов библиотеки API JavaScript в проекте до нового выпуска
В следующих действиях Office.js файлы библиотеки до последней версии. Эти действия используются Visual Studio 2019 г., но они аналогичны для предыдущих версий Visual Studio.

1. В Visual Studio 2019 г. откройте или создайте новый **проект Office надстройки.**
2. Выберите   >  **инструменты NuGet диспетчер пакетов**  >  **Управление пакетами Nuget для решения**.
3. Выберите вкладку **Обновления**.
4. Выберите Microsoft.Office.js. Убедитесь, что источник **пакета nuget.org**.
5. В левой области выберите **Установите и** завершите процесс обновления пакета.

Вам потребуется выполнить несколько дополнительных действий, чтобы завершить обновление. В  теге head of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated Office JavaScript API follows:

  ```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
  ```

   > [!NOTE] 
   > `/1/` в `office.js` в URL-адресе CDN указывает на то, что необходимо использовать последний добавочный выпуск Office.js версии 1.


### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>Обновление схемы манифеста в проекте до версии 1.1

В файле манифеста надстройки обновите атрибут **xmlns** элемента **OfficeApp**, заменив значение версии на `1.1` и оставив все атрибуты, кроме **xmlns**, без изменений.

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> После обновления версии схемы манифеста надстройки до 1.1 необходимо удалить элементы **Capabilities** и **Capability** и заменить их элементами [Hosts](../reference/manifest/hosts.md) и [Host](../reference/manifest/host.md) или элементами Requirements [and Requirements](specify-office-hosts-and-api-requirements.md).

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a>Обновление проекта надстройки Office, созданного с помощью текстового редактора или другой среды IDE

Для проектов, созданных до выпуска v1.1 API Office JavaScript и схемы манифеста надстройки, необходимо обновить HTML-страницы надстройки для ссылки CDN библиотеки v1.1, а также обновить манифестный файл надстройки, чтобы использовать схему v1.1. 

Процесс обновления применяется к _проектам по отдельности_. Вам потребуется повторить его для каждого проекта надстройки, в котором вы хотите использовать файл Office.js и схему манифеста надстройки версии 1.1.

Вам не нужны локальные копии файлов API Office JavaScript (Office.js и файлов .js приложений) для разработки надстройкиOffice (ссылаясь на CDN для Office.js загрузки необходимых файлов во время работы), но если вы хотите локализованную копию файлов библиотеки, вы можете использовать утилиту [NuGet Command-Line](https://docs.nuget.org/consume/installing-nuget) и команду для их `Install-Package Microsoft.Office.js` скачивания.

> [!NOTE]
> Чтобы получить копию XSD (определения схемы XML) для манифеста надстройки версии 1.1, см. статью [Справочник по схеме манифестов надстроек Office (версия 1.1)](../develop/add-in-manifests.md).


### <a name="update-the-office-javascript-api-library-files-in-your-project-to-use-the-newest-release"></a>Обновление Office файлов библиотеки API JavaScript в проекте, чтобы использовать самый новый выпуск

1. Откройте HTML-страницы надстройки в текстовом редакторе или интегрированной среде разработки.

2. В  теге head of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated Office JavaScript API follows:

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
    ```

   > [!NOTE]
   > `/1/` перед `office.js` в URL-адресе CDN указывает на то, что необходимо использовать последний добавочный выпуск Office.js версии 1.

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>Обновление схемы манифеста в проекте до версии 1.1

В файле манифеста надстройки обновите атрибут **xmlns** элемента **OfficeApp**, заменив значение версии на `1.1` и оставив все атрибуты, кроме **xmlns**, без изменений.

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> После обновления версии схемы манифеста надстройки до 1.1 необходимо удалить элементы **Capabilities** и **Capability** и заменить их элементами [Hosts](../reference/manifest/hosts.md) и [Host](../reference/manifest/host.md) или элементами Requirements [and Requirements](specify-office-hosts-and-api-requirements.md).

## <a name="see-also"></a>См. также

- [Укажите Office приложения и требования к API](specify-office-hosts-and-api-requirements.md) ]
- [Общие сведения об API JavaScript для Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript для Office](../reference/javascript-api-for-office.md)
- [Справка по схеме для манифестов надстроек Office (версия 1.1)](../develop/add-in-manifests.md)
