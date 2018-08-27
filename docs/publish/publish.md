---
title: Развертывание и публикация надстройки Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 83581b729f5004c36d267bda14795275a5153a9c
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925348"
---
# <a name="deploy-and-publish-your-office-add-in"></a>Развертывание и публикация надстройки Office

Для тестирования или распространения надстройки Office можно использовать один из указанных ниже способов.

|**Способ**|**Применение**|
|:---------|:------------|
|[Загрузка неопубликованного приложения](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|При разработке для проверки работы надстройки в Windows, Office Online, на iPad или Mac.|
|[Централизованное развертывание](centralized-deployment.md)|В облачном или гибридном развертывании для распространения надстройки в организации с помощью Центра администрирования Office 365.|
|[Каталог SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|В локальной среде для распространения надстройки в организации.|
|[AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)|Для распространения надстройки среди всех пользователей.|
|[Сервер Exchange Server](#outlook-add-in-deployment)|В локальной или облачной среде для распространения надстроек Outlook.|
|[Общая сетевая папка](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|На том расположенном в сети компьютере с Windows, где должна размещаться надстройка, перейдите к родительской папке или диску с папкой, которую требуется использовать в качестве каталога общих папок.|

> [!NOTE]
> Если вы планируете [опубликовать](../publish/publish.md) надстройку в AppSource и сделать ее доступной в интерфейсе Office, убедитесь, что она соответствует [политикам проверки AppSource](https://docs.microsoft.com/office/dev/store/validation-policies). Например, чтобы пройти проверку, надстройка должна работать на всех платформах, поддерживающих определенные вами методы. Дополнительные сведения см. в [разделе 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) и на [странице со сведениями о доступности и ведущих приложениях для надстроек Office](../overview/office-add-in-availability.md).

## <a name="deployment-options-by-office-host"></a>Варианты развертывания для различных ведущих приложений Office

Доступные варианты развертывания зависят от ведущего приложения Office и типа создаваемой надстройки.

### <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a>Варианты развертывания для надстроек Word, Excel и PowerPoint

| Точка расширения | Загрузка неопубликованного приложения | Центр администрирования Office 365 |AppSource| Каталог SharePoint\*  |
|:----------------|:-----------:|:-----------------------:|:----------:|:--------------------:|
| Контент         | X           | X                       | X          | X                    |
| Область задач       | X           | X                       | X          | X                    |
| Команда           | X           | X                       | X          |                      |

* Каталоги SharePoint не поддерживают Office 2016 для Mac.

### <a name="deployment-options-for-outlook-add-ins"></a>Способы развертывания надстроек Outlook

| Точка расширения | Загрузка неопубликованного приложения | Сервер Exchange Server | AppSource |
|:----------------|:-----------:|:---------------:|:------------:|
| Почтовое приложение        | X           | X               | X            |
| Команда         | X           | X               | X            |

## <a name="deployment-methods"></a>Методы развертывания

Указанные ниже разделы содержат дополнительные сведения о методах развертывания, которые чаще всего используются для распространения надстроек Office в организации.

Сведения о том, как пользователи приобретают, вставляют и запускают надстройки, см. в статье [Начало работы с надстройкой Office](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE).

### <a name="centralized-deployment-via-the-office-365-admin-center"></a>Централизованное развертывание в Центре администрирования Office 365 

С помощью Центра администрирования Office 365 администраторы могут с легкостью развертывать надстройки Office для пользователей и групп в организации. В этом случае надстройки становятся доступны в приложениях Office сразу. Настраивать клиенты не требуется. Используя централизованное развертывание, можно развертывать как внутренние надстройки, так и те, что предоставляются независимыми поставщиками программного обеспечения.

Дополнительные сведения см. в разделе [Публикация надстроек Office с помощью централизованного развертывания в Центре администрирования Office 365](centralized-deployment.md).

### <a name="sharepoint-catalog-deployment"></a>Развертывание с использованием каталога SharePoint

Каталог надстроек SharePoint — это специальный семейство веб-сайтов, в котором можно размещать надстройки Word, Excel и PowerPoint. Та как каталоги SharePoint не поддерживают новые функции надстроек, реализованные в узле `VersionOverrides` манифеста, в том числе команды надстроек, рекомендуем развертывать надстройки в Центре администрирования. Команды надстроек, развернутые с помощью каталога SharePoint, по умолчанию открываются в области задач.

Если вы развертываете надстройки в локальной среде, используйте каталог SharePoint. Дополнительные сведения см. в статье [Публикация надстроек области задач и контентных надстроек в каталоге SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

> [!NOTE]
> Каталоги SharePoint не поддерживают Office 2016 для Mac. Чтобы развернуть надстройки Office на клиентах Mac, их необходимо отправить в [AppSource]. 

### <a name="outlook-add-in-deployment"></a>Развертывание надстроек Outlook

В локальных и онлайн-средах, в которых не используется служба идентификации Azure AD, надстройки Outlook можно развертывать через сервер Exchange Server. 

Для развертывания надстроек Outlook требуется следующее:

- Office 365, Exchange Online или Exchange Server 2013 или более поздней версии
- Outlook 2013 или более поздней версии

Чтобы назначить надстройки клиентам, загрузите манифест напрямую из файла или URL-адреса в Центре администрирования Exchange или добавьте надстройку из AppSource. Чтобы назначить надстройки отдельным пользователям, необходимо использовать Exchange PowerShell. Дополнительные сведения см. в статье [Установка или удаление надстроек Outlook для организации](https://technet.microsoft.com/library/jj943752(v=exchg.150).aspx) на сайте TechNet.

## <a name="see-also"></a>См. также

- [Загрузка неопубликованных надстроек Outlook для тестирования](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- [Отправка в AppSource][AppSource]
- [Рекомендации по проектированию надстроек Office](../design/add-in-design.md)
- [Создание эффективных описаний в AppSource](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings)
- [Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office](../testing/testing-and-troubleshooting.md)

[AppSource]: https://docs.microsoft.com/office/dev/store/submit-to-the-office-store
[Office Add-in host and platform availability]: ../overview/office-add-in-availability
