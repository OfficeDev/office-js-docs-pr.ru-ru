---
title: Требования для запуска надстроек Office
description: Узнайте о требованиях клиента и сервера, которые необходимо выполнить конечному пользователю Office надстройки.
ms.date: 07/27/2021
localization_priority: Normal
ms.openlocfilehash: 1cc591db443c1fb0e2ca934cd05f52ad41ed61cf977ef4053af70d536867a6db
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57082992"
---
# <a name="requirements-for-running-office-add-ins"></a>Требования для запуска надстроек Office

В этой статье описаны требования к программному обеспечению и устройствам для запуска надстроек Office.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

Чтобы узнать, где поддерживаются Office, см. в Office клиентского приложения и доступности платформы для [Office надстройки.](../overview/office-add-in-availability.md)

## <a name="server-requirements"></a>Требования к серверу

Чтобы иметь возможность установить и запустить любую Надстройка Office, необходимо сначала развернуть файлы манифеста и веб-страниц для пользовательского интерфейса и кода надстройки в соответствующих папках на сервере.

Для всех типов надстроек (контентных надстроек, надстроек Outlook и надстроек области задач, а также команд надстроек) необходимо развертывать файлы веб-страниц на веб-сервере или в службе веб-хостинга, например [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md).

[!include[HTTPS guidance](../includes/https-guidance.md)]

> [!TIP]
> Когда вы создаете и отлаживаете надстройку в Visual Studio, эта система развертывает и запускает соответствующие файлы веб-страниц локально с помощью IIS Express. Использовать дополнительный веб-сервер не требуется.

Для надстройок контента и области задач в поддерживаемых Office клиентских приложениях — Excel, PowerPoint, Project или Word [](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) — также требуется каталог приложений в SharePoint для загрузки XML-файла манифеста надстройки или необходимо развернуть надстройку с помощью интегрированных [приложений](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps).

Чтобы проверить и запустить надстройку Outlook, учетная запись электронной почты Outlook пользователя должна находиться в Exchange 2013 г. или более поздней, которая доступна через Microsoft 365, Exchange Online или с помощью локальной установки. Пользователь или администратор устанавливают файлы манифестов надстроек Outlook на соответствующем сервере.

> [!NOTE]
> Учетные записи POP и IMAP в Outlook не поддерживают надстройки Office.

## <a name="client-requirements-windows-desktop-and-tablet"></a>Требования к клиенту: компьютеры и планшеты под управлением Windows

Следующее программное обеспечение необходимо для разработки надстройки Office для поддерживаемых Office или веб-клиентов, которые работают на Windows настольных, ноутбуках или планшетных устройствах.

- Для настольных компьютеров под управлением 32- и 64-разрядных версий Windows, а также таких планшетов, как Surface Pro:
  - 32- или 64-разрядная версия Office 2013 или более поздняя версия в Windows 7 или более поздней версии.
  - Excel 2013, Outlook 2013, PowerPoint 2013, Project профессиональный 2013, Project 2013 с пакетом обновления 1 (SP1), Word 2013 или более поздняя версия клиента Office, если надстройка Office тестируется или запускается специально для одного из этих клиентов Office. Клиенты Office для настольных ПК можно устанавливать локально или на клиентском компьютере с помощью технологии "нажми и работай".

  Если у вас есть Microsoft 365 подписка и у вас нет доступа к клиенту Office, вы можете скачать и установить последнюю версию [Office](https://support.office.com/article/download-and-install-or-reinstall-office-365-or-office-2019-on-a-pc-or-mac-4414eaaf-0478-48be-9c42-23adc4716658).

- Браузер Internet Explorer 11 или Microsoft Edge (в зависимости от версий Windows и Office) должен быть установлен, но может не использоваться по умолчанию. Для поддержки надстроек Office клиент Office, выступающий в роли ведущего приложения, использует компоненты браузера, которые входят в состав Internet Explorer 11 или Microsoft Edge. Дополнительные сведения см. в статье [Браузеры, используемые надстройками Office](browsers-used-by-office-web-add-ins.md).

  > [!NOTE]
  > Для работы веб-надстроек Office необходимо отключить конфигурацию усиленной безопасности Internet Explorer (ESC). Если вы используете компьютер с Windows Server в качестве клиента при разработке надстроек, учитывайте, что конфигурация ESC включена по умолчанию в Windows Server.

- По умолчанию используется один из следующих браузеров: Internet Explorer 11 или последняя версия Microsoft Edge, Chrome, Firefox или Safari (Mac OS).
- Редактор HTML и JavaScript, например "Блокнот", [Visual Studio и Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs) или стороннее средство веб-разработки.

## <a name="client-requirements-os-x-desktop"></a>Требования к клиенту: настольный компьютер OS X

Outlook Mac, который распространяется в Microsoft 365, поддерживает Outlook надстройки. Запуск Outlook в Outlook Mac имеет те же требования, что и Outlook на самом Mac: операционная система должна быть как минимум OS X v10.10 "Yosemite". Outlook для Mac использует WebKit в качестве обработчика макетов для преобразования страниц надстройки, поэтому дополнительные зависимости от браузеров отсутствуют.

Ниже приведены минимальные версии клиентов Office для Mac, которые поддерживают надстройки Office.

- Word версии 15.18 (160109)
- Excel версии 15.19 (160206)
- PowerPoint версии 15.24 (160614)

## <a name="client-requirements-browser-support-for-office-web-clients-and-sharepoint"></a>Требования к клиенту: поддержка браузеров для веб-клиентов Office в Интернете и SharePoint

Любой браузер, поддерживающий ECMAScript 5.1, HTML5 и CSS3, например Internet Explorer 11 либо последняя версия Microsoft Edge, Chrome, Firefox или Safari (Mac OS).


## <a name="client-requirements-non-windows-smartphone-and-tablet"></a>Требования к клиенту: смартфоны и планшеты под управлением операционных систем, отличных от Windows

Специально для приложения Outlook, работающего в браузере на смартфонах и планшетах под управлением систем, отличных от Windows, для тестирования и запуска надстроек Outlook необходимо следующее программное обеспечение:


| Приложение Office | Устройство | Операционная система | Учетная запись Exchange | Мобильный браузер |
|:-----|:-----|:-----|:-----|:-----|
|Outlook для Android|Планшеты и смартфоны с Android|Android 4.4 KitKat или более поздней версии|О последнем обновлении Приложения Microsoft 365 для бизнеса или Exchange Online|Встроенное приложение для Android, браузер не применим|
|Outlook для iOS|Планшеты iPad, смартфоны iPhone|iOS 11 или более поздняя версия|О последнем обновлении Приложения Microsoft 365 для бизнеса или Exchange Online|Встроенное приложение для iOS, браузер не применим|
|Outlook в Интернете|iPhone 4, iPad 2, iPod Touch 4 или более поздние модели этих устройств|iOS 5 или более поздняя версия|В Microsoft 365, Exchange Online или в Exchange Server 2013 г. или более поздней|Safari|

> [!NOTE]
> Встроенные приложения OWA для Android, OWA для iPad и OWA для iPhone [устарели](https://support.office.com/article/Microsoft-OWA-mobile-apps-are-being-retired-076ec122-4576-4900-bc26-937f84d25a4b) и больше не требуются и не применяются для тестирования надстроек Outlook.


## <a name="see-also"></a>См. также

- [Обзор платформы надстроек Office](../overview/office-add-ins.md)
- [Доступность клиентских приложений и платформ Office для надстроек Office](../overview/office-add-in-availability.md)
- [Браузеры, используемые надстройками Office](browsers-used-by-office-web-add-ins.md)
