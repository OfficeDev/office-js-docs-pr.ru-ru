---
title: Развертывание и публикация надстроек Office
description: Методы и варианты развертывания надстройки Office для тестирования и распространения.
ms.date: 07/30/2021
localization_priority: Priority
ms.openlocfilehash: 28589d71d7b7e59640ce11fe231671ca2b3c65fb
ms.sourcegitcommit: 69f6492de8a4c91e734250c76681c44b3f349440
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/03/2021
ms.locfileid: "58868717"
---
# <a name="deploy-and-publish-office-add-ins"></a>Развертывание и публикация надстроек Office

Для тестирования или распространения надстройки Office можно использовать один из указанных ниже способов.

|**Способ**|**Применение**|
|:---------|:------------|
|[Загрузка неопубликованного приложения](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing)|При разработке для проверки работы надстройки в Windows, iPad, Mac или в браузере. (Не для введенных в эксплуатацию надстроек).|
|[Общая сетевая папка](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|Для тестирования в процессе разработки надстройки, работающей в Windows, после ее публикации на сервере, отличном от localhost. (Не для типовых надстроек и не для тестирования на iPad, Mac или в Интернете).|
|[AppSource](/office/dev/store/submit-to-appsource-via-partner-center)|Для распространения надстройки среди всех пользователей.|
|[Центр администрирования Microsoft 365](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps)|Для распространения надстройки среди пользователей в организации с помощью Центра администрирования Microsoft 365 в процессе облачного развертывания. Реализуется с помощью [интегрированных приложений](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps) или [централизованного развертывания](/microsoft-365/admin/manage/centralized-deployment-of-add-ins). |
|[Каталог SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|В локальной среде для распространения надстройки в организации.|
|[Сервер Exchange Server](#outlook-add-in-deployment)|В локальной или облачной среде для распространения надстроек Outlook.|

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="deployment-options-by-office-application-and-add-in-type"></a>Варианты развертывания по типам приложений и надстроек Office

Доступные варианты развертывания зависят от приложения Office и типа создаваемой надстройки.

### <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a>Варианты развертывания для надстроек Word, Excel и PowerPoint

| Точка расширения | Загрузка неопубликованного приложения | Общая сетевая папка | AppSource | Центр администрирования Microsoft 365 | Каталог SharePoint\* |
|:----------------|:-----------:|:-------------:|:---------:|:--------------------------:|:--------------------:|
| Контент         | X           | X             | X         | X                          | X                    |
| Область задач       | X           | X             | X         | X                          | X                    |
| Команда         | X           | X             | X         | X                          |                      |

&#42; Каталоги SharePoint не поддерживают Office для Mac.

### <a name="deployment-options-for-outlook-add-ins"></a>Варианты развертывания надстроек Outlook

| Точка расширения | Загрузка неопубликованного приложения | AppSource | Сервер Exchange Server |
|:----------------|:-----------:|:---------:|:---------------:|
| Почтовое приложение        | X           | X         | X               |
| Команда         | X           | X         | X               |

## <a name="production-deployment-methods"></a>Методы развертывания в рабочей среде

Указанные ниже разделы содержат дополнительные сведения о методах развертывания, которые чаще всего используются для распространения типовых надстроек Office среди пользователей в организации.

Сведения о том, как пользователи получают, устанавливают и запускают надстройки, см. в статье [Начало работы с надстройкой Office](https://support.microsoft.com/office/82e665c4-6700-4b56-a3f3-ef5441996862).

### <a name="integrated-apps-via-the-microsoft-365-admin-center"></a>Интегрированные приложения в Центре администрирования Microsoft 365

Центр администрирования Microsoft 365 позволяет администраторам легко развертывать надстройки Office для пользователей и групп в организации. При развертывании с помощью Центра администрирования надстройки становятся доступны в приложениях Office немедленно. Настраивать клиенты не требуется. Используя интегрированные приложения, можно распространять как внутренние надстройки, так и те, что предоставляются независимыми поставщиками программного обеспечения (ISV). Кроме того, в интегрированных приложениях отображаются надстройки для администраторов и другие приложения, связанные в единый пакет одним и тем же ISV, что позволяет использовать их на всей платформе Microsoft 365.

Объединяя свои надстройки Office, приложения Teams, приложения SPFx и [другие приложения](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps#what-apps-can-i-deploy-from-integrated-apps), вы создаете для своих клиентов единое предложение пакета программного обеспечения как услуги (SaaS). Общую информацию об этом процессе см. в статье [Как составить план предложения SaaS для Commercial Marketplace](/azure/marketplace/plan-saas-offer). Подробные сведения о создании интегрированных приложений см. в статье [Настройка интеграции приложений Microsoft 365](/azure/marketplace/create-new-saas-offer#configure-microsoft-365-app-integration).

Подробнее о процессе развертывания интегрированных приложений см. в статье [Тестирование и развертывание приложений Microsoft 365 партнерами на портале интегрированных приложений](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps).

> [!IMPORTANT]
> У клиентов, работающих в изолированных или государственных облаках, нет доступа к интегрированным приложениям. Вместо этого они применяют централизованное развертывание. Централизованное развертывание — похожий метод развертывания, но при его использовании связанные надстройки и приложения не представляются администратору. Для получения дополнительных сведений см. статью [Как определить, подходит ли централизованное развертывание надстроек для вашей организации](/microsoft-365/admin/manage/centralized-deployment-of-add-ins).

### <a name="sharepoint-app-catalog-deployment"></a>Развертывание с использованием каталога приложений SharePoint

Каталог приложений SharePoint — это специальное семейство веб-сайтов, в котором можно размещать надстройки Word, Excel и PowerPoint. Так как каталоги SharePoint не поддерживают новые функции надстроек, реализованные в узле `VersionOverrides` манифеста, в том числе команды надстроек, рекомендуем развертывать надстройки в Центре администрирования. Команды надстроек, развернутые с помощью каталога SharePoint, по умолчанию открываются в области задач.

Если вы развертываете надстройки в локальной среде, используйте каталог SharePoint. Дополнительные сведения см. в статье [Публикация надстроек области задач и контентных надстроек в каталоге SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

> [!NOTE]
> Каталоги SharePoint не поддерживают Office для Mac. Для развертывания надстроек Office на клиентах Mac необходимо отправить их в [AppSource](/office/dev/store/submit-to-the-office-store).

### <a name="outlook-add-in-deployment"></a>Развертывание надстроек Outlook

В локальных и онлайн-средах, в которых не используется служба идентификации Azure AD, надстройки Outlook можно развертывать через сервер Exchange Server.

Для развертывания надстроек Outlook требуется следующее:

- Microsoft 365, Exchange Online или Exchange Server 2013 или более поздней версии
- Outlook 2013 или более поздней версии

Чтобы назначить надстройки клиентам, загрузите манифест напрямую из файла или URL-адреса в Центре администрирования Exchange или добавьте надстройку из AppSource. Чтобы назначить надстройки отдельным пользователям, необходимо использовать Exchange PowerShell. Дополнительные сведения см. в статье [Установка или удаление надстроек Outlook для организации](/exchange/clients-and-mobile-in-exchange-online/add-ins-for-outlook/install-or-remove-outlook-add-ins) на сайте TechNet.

## <a name="see-also"></a>См. также

- [Загрузка неопубликованных надстроек Outlook для тестирования](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- [Отправка в AppSource][AppSource]
- [Рекомендации по разработке надстроек Office](../design/add-in-design.md)
- [Создание эффективных описаний в AppSource](/office/dev/store/create-effective-office-store-listings)
- [Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office](../testing/testing-and-troubleshooting.md)
- [Что такое Microsoft Commercial Marketplace?](/azure/marketplace/overview)

[AppSource]: /office/dev/store/submit-to-appsource-via-partner-center
