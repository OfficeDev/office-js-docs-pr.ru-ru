---
title: Требования к надстройкам Outlook
description: Чтобы надстройки Outlook загружались и работали надлежащим образом, существует ряд требований к серверам и клиентам.
ms.date: 02/09/2021
ms.localizationpriority: high
ms.openlocfilehash: 0946cf6841fd46957b21b0d24698042e8d3fcb8c
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746201"
---
# <a name="outlook-add-in-requirements"></a>Требования к надстройкам Outlook

Чтобы надстройки Outlook загружались и работали надлежащим образом, существует ряд требований к серверам и клиентам.

## <a name="client-requirements"></a>Требования к клиентам

- Клиент должен быть одним из поддерживаемых приложений для надстроек Outlook. Эти клиенты поддерживают надстройки.

  - Outlook 2013 или более поздней версии для Windows
  - Outlook 2016 или более поздней версии для Mac
  - Outlook для iOS
  - Outlook для Android
  - Outlook в Интернете для Exchange 2016 или более поздней версии
  - Outlook в Интернете для Exchange 2013
  - Outlook.com.

- Клиент должен напрямую подключаться к серверу Exchange Server или Microsoft 365. При настройке клиента пользователь должен выбрать тип учетной записи **Exchange**, **Office** или **Outlook.com**. Если клиент настроен на подключение POP3 или IMAP, надстройки не загрузятся.

## <a name="mail-server-requirements"></a>Требования к почтовым серверам

Если пользователь подключен к Microsoft 365 или Outlook.com, требования к почтовому серверу уже выполнены. Но если пользователи подключаются к локально установленным экземплярам Exchange Server, требуется соответствие указанным ниже условиям.

- Должен использоваться сервер Exchange 2013 или более поздней версии.
- Веб-службы Exchange (EWS) должны быть включены и подключены к Интернету. Многие надстройки требуют надлежащей работы EWS.
- Чтобы сервер мог издавать действительные маркеры идентификации, он должен иметь действительный сертификат проверки подлинности. Новые установленные экземпляры сервера Exchange Server обладают сертификатом проверки подлинности по умолчанию. Дополнительные сведения см. в статьях [Цифровые сертификаты и шифрование в Exchange 2016](/Exchange/architecture/client-access/certificates) и [Set-AuthConfig](/powershell/module/exchange/organization/Set-AuthConfig).
- Для получения доступа к надстройкам из [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=a35323d5-0e3d-4cc0-ba44-57537d74aae8&omexanonuid=581941df-1c6f-4eda-89e7-651af8aeaeb2) серверы клиентского доступа должны быть настроены на связь с AppSource.

## <a name="add-in-server-requirements"></a>Требования к серверам надстроек

Файлы надстройки (например, HTML, JavaScript) могут быть размещены на любой платформе веб-сервера. Единственное требование — настройка сервера на использование HTTPS и доверия к SSL-сертификату со стороны клиента.

## <a name="see-also"></a>См. также

- [Требования для запуска надстроек Office](../concepts/requirements-for-running-office-add-ins.md)
- [Доступность клиентских приложений и платформ Office для надстроек Office (раздел Outlook)](../overview/office-add-in-availability.md#outlook)
- [Поддержка наборов обязательных элементов API JavaScript для Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)
