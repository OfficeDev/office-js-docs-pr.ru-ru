---
title: Обновление схемы манифеста надстройки API JavaScript для Office и версии 1,1 до последней версии
description: Обновление до версии 1.1 файлов JavaScript (Office.js и JS-файлов приложения) и файла проверки манифеста надстройки в проекте надстройки Office.
ms.date: 10/11/2019
localization_priority: Normal
ms.openlocfilehash: 34127b3920af1309d4e4c2e1c265c676640a1c24
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093555"
---
# <a name="update-to-the-latest-office-javascript-api-library-and-version-11-add-in-manifest-schema"></a>Обновление схемы манифеста надстройки API JavaScript для Office и версии 1,1 до последней версии

В этой статье рассказывается, как обновить файлы JavaScript (Office.js и JS-файлы для конкретной надстройки) и файл проверки манифеста надстройки в проекте надстройки Office до версии 1.1.

> [!NOTE]
> Проекты, созданные в Visual Studio 2019, уже используют версию 1,1. Однако для версии 1.1 периодически выпускаются незначительные обновления, которые можно применить с помощью методов, описанных в этой статье.

## <a name="use-the-most-up-to-date-project-files"></a>Использование последних версий файлов в проекте

Если вы используете Visual Studio для разработки надстройки, для использования новейших элементов API JavaScript для Office и [функций версии 1.1 манифеста надстройки](../develop/add-in-manifests.md) (которая проверяется по сравнению с offappmanifest-1.1. xsd) необходимо скачать Visual Studio 2019. Чтобы скачать Visual Studio 2019, посетите [страницу Visual Studio IDE](https://visualstudio.microsoft.com/vs/). Во время установки потребуется выбрать рабочую нагрузку разработки Office и SharePoint.

Если вы используете текстовый редактор или другую интегрированную среду разработки, отличную от Visual Studio, чтобы разработать надстройка, обновите ссылки на CDN для файла Office.js и версию схемы, на которую ссылается манифест приложения для Office.

Для запуска надстройки, разработанной с помощью новых и обновленных функций Office.js API и манифеста надстроек, пользователям должны быть назначены продукты Office 2013 с пакетом обновления 1 (SP1) или более поздней версии, а также в случаях, когда это возможно, SharePoint Server 2013 SP1 и родственных серверных продуктов, Exchange Server 2013 с пакетом обновления 1 (SP1) или эквивалентных сетевых 365 продуктов

Сведения о том, как скачать Office, SharePoint и Exchange с пакетом обновления 1, см. в следующих статьях:

- [Список всех пакетов обновления 1 (SP1) для Microsoft Office 2013 и связанных продуктов для настольных систем](https://support.microsoft.com/kb/2850036)

- [Список всех пакетов обновления 1 (SP1) для Microsoft SharePoint Server 2013 и связанных серверных продуктов](https://support.microsoft.com/kb/2850035)

- [Описание пакета обновления 1 для Exchange Server 2013](https://support.microsoft.com/kb/2926248)


## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a>Обновление проекта надстройки Office, созданного в Visual Studio

Для проектов, созданных до выпуска версии 1.1 API JavaScript для Office и схемы манифеста надстройки, можно обновить файлы проекта с помощью **диспетчера пакетов NuGet**, а затем обновить HTML-страницы надстройки, чтобы они ссылались на них. 

Обратите внимание, что процесс обновления применяется к _проектам по отдельности_. Вам потребуется повторить его для каждого проекта надстройки, в котором вы хотите использовать Office.js и схему манифеста надстройки версии 1.1.

### <a name="update-the-office-javascript-api-library-files-in-your-project-to-the-newest-release"></a>Обновление файлов библиотеки API JavaScript для Office в проекте до последней версии
Приведенные ниже действия приведут к обновлению файлов библиотеки Office.js до последней версии. В этом разделе описано, как использовать Visual Studio 2019, но они аналогичны предыдущим версиям Visual Studio.

1. В Visual Studio 2019 откройте или создайте новый проект **надстройки Office** .
2. Выберите **инструменты**  >  **Диспетчер пакетов NuGet**  >  **Управление пакетами NuGet для решения**.
3. Выберите вкладку **Обновления**.
4. Выберите Microsoft.Office.js. Убедитесь, что источник пакета находится в **NuGet.org**.
5. В левой области выберите **установить** и завершить процесс обновления пакета.

Вам потребуется выполнить несколько дополнительных действий, чтобы завершить обновление. В теге **head** HTML-страниц надстройки закомментируйте или удалите существующие ссылки на скрипты office.js, а также ссылку на обновленную библиотеку API JavaScript для Office следующим образом:

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
> После обновления версии схемы манифеста надстройки до 1,1 необходимо удалить элементы **capabilities** и **capability** и заменить их на элементы [hosts](../reference/manifest/hosts.md) и [Host](../reference/manifest/host.md) , а также [элементы требований и требований](specify-office-hosts-and-api-requirements.md).

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a>Обновление проекта надстройки Office, созданного с помощью текстового редактора или другой среды IDE

Для проектов, созданных до выпуска версии 1.1 API JavaScript для Office и схемы манифеста надстройки, необходимо обновить HTML-страницы надстройки, чтобы ссылаться на сеть CDN библиотеки версии 1.1, и обновить файл манифеста надстройки для использования схемы версии 1.1. 

Процесс обновления применяется к _проектам по отдельности_. Вам потребуется повторить его для каждого проекта надстройки, в котором вы хотите использовать файл Office.js и схему манифеста надстройки версии 1.1.

Локальные копии файлов API JavaScript для Office (Office.js и JS-файлов приложения) не требуются для разработки надстройки подпиской (ссылка на CDN для Office.js загружает необходимые файлы во время выполнения), но если вам нужна локальная копия файлов библиотеки, вы можете использовать [служебную программу командной строки NuGet](https://docs.nuget.org/consume/installing-nuget) и `Install-Package Microsoft.Office.js` команду для их загрузки.

> [!NOTE]
> Чтобы получить копию XSD (определения схемы XML) для манифеста надстройки версии 1.1, см. статью [Справочник по схеме манифестов надстроек Office (версия 1.1)](../develop/add-in-manifests.md).


### <a name="update-the-office-javascript-api-library-files-in-your-project-to-use-the-newest-release"></a>Обновление файлов библиотеки API JavaScript для Office в проекте для использования последнего выпуска

1. Откройте HTML-страницы надстройки в текстовом редакторе или интегрированной среде разработки.

2. В теге **head** HTML-страниц надстройки закомментируйте или удалите существующие ссылки на скрипты office.js, а также ссылку на обновленную библиотеку API JavaScript для Office следующим образом:

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
> После обновления версии схемы манифеста надстройки до 1,1 необходимо удалить элементы **capabilities** и **capability** и заменить их на элементы [hosts](../reference/manifest/hosts.md) и [Host](../reference/manifest/host.md) , а также [элементы требований и требований](specify-office-hosts-and-api-requirements.md).

## <a name="see-also"></a>См. также

- [Указание ведущих приложений Office и требований к API](specify-office-hosts-and-api-requirements.md) ]
- [Общие сведения об API JavaScript для Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript для Office](../reference/javascript-api-for-office.md)
- [Справка по схеме для манифестов надстроек Office (версия 1.1)](../develop/add-in-manifests.md)
