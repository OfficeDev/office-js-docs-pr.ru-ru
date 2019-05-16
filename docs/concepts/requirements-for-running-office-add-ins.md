---
title: Требования для запуска надстроек Office
description: ''
ms.date: 05/14/2019
localization_priority: Priority
ms.openlocfilehash: 2dcdfb2562233550016cd2d04571239318ffffa3
ms.sourcegitcommit: 944cbb5c6ce055f6db1833182b24d490d1dce01d
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/14/2019
ms.locfileid: "33992184"
---
# <a name="requirements-for-running-office-add-ins"></a>Требования для запуска надстроек Office

В этой статье описаны требования к программному обеспечению и устройствам для запуска надстроек Office.

> [!NOTE]
> Если вы планируете [опубликовать](../publish/publish.md) надстройку в AppSource и сделать ее доступной в интерфейсе Office, убедитесь, что она соответствует [политикам проверки AppSource](/office/dev/store/validation-policies). Например, чтобы пройти проверку, надстройка должна работать на всех платформах, поддерживающих определенные вами методы. Дополнительные сведения см. в [разделе 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) и на [странице со сведениями о доступности и ведущих приложениях для надстроек Office](../overview/office-add-in-availability.md).

Общие сведения о том, на каких платформах поддерживаются надстройки Office, см. в статье [Доступность ведущих приложений и платформ для надстроек Office](../overview/office-add-in-availability.md).

## <a name="server-requirements"></a>Требования к серверу

Чтобы иметь возможность установить и запустить любую Надстройка Office, необходимо сначала развернуть файлы манифеста и веб-страниц для пользовательского интерфейса и кода надстройки в соответствующих папках на сервере.

Для всех типов надстроек (контентных надстроек, надстроек Outlook и надстроек области задач, а также команд надстроек) необходимо развертывать файлы веб-страниц на веб-сервере или в службе веб-хостинга, например [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md).

[!include[HTTPS guidance](../includes/https-guidance.md)]

> [!TIP]
> Когда вы создаете и отлаживаете надстройку в Visual Studio, эта система развертывает и запускает соответствующие файлы веб-страниц локально с помощью IIS Express. Использовать дополнительный веб-сервер не требуется.

Кроме того, требуется [каталог надстроек](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) в SharePoint, чтобы отправить XML-файл манифеста надстройки (контентной или области задач) в поддерживаемых ведущих приложениях Office — веб-приложениях Access, Word, Excel, PowerPoint и Project.

Чтобы тестировать и запускать надстройки Outlook, необходимо разместить учетную запись электронной почты Outlook в Exchange 2013 или более поздней версии, доступ к которой можно получить в Office 365, Exchange Online или в локально установленной версии. Пользователь или администратор устанавливают файлы манифестов надстроек Outlook на соответствующем сервере.

> [!NOTE]
> Учетные записи POP и IMAP в Outlook не поддерживают надстройки Office.

## <a name="client-requirements-windows-desktop-and-tablet"></a>Требования к клиенту: компьютеры и планшеты под управлением Windows

Чтобы можно было разработать Надстройка Office для поддерживаемых классических клиентов Office или веб-клиентов, работающих на настольных компьютерах, ноутбуках или планшетах с ОС Windows, необходимо следующее программное обеспечение:


- Для настольных компьютеров под управлением 32- и 64-разрядных версий Windows, а также таких планшетов, как Surface Pro:
    - 32- или 64-разрядная версия Office 2013 или более поздняя версия в Windows 7 или более поздней версии.
    - Excel 2013, Outlook 2013, PowerPoint 2013, Project профессиональный 2013, Project 2013 с пакетом обновления 1 (SP1), Word 2013 или более поздняя версия клиента Office, если надстройка Office тестируется или запускается специально для одного из этих клиентов Office. Клиенты Office для настольных ПК можно устанавливать локально или на клиентском компьютере с помощью технологии "нажми и работай".

  Если у вас не установлен клиент Office, но есть подписка на Office 365, вы можете [скачать и установить последнюю версию Office](https://support.office.com/article/download-and-install-or-reinstall-office-365-or-office-2019-on-a-pc-or-mac-4414eaaf-0478-48be-9c42-23adc4716658).

- Браузер Internet Explorer 11 или Microsoft Edge (в зависимости от версий Windows и Office) должен быть установлен, но может не использоваться по умолчанию. Для поддержки надстроек Office клиент Office, выступающий в роли ведущего приложения, использует компоненты браузера, которые входят в состав Internet Explorer 11 или Microsoft Edge. Дополнительные сведения см. в статье [Веб-средства просмотра, используемые надстройками Office](web-viewers-used-by-office-web-add-ins.md).

  > [!NOTE]
  > Для работы веб-надстроек Office необходимо отключить конфигурацию усиленной безопасности Internet Explorer (ESC). Если вы используете компьютер с Windows Server в качестве клиента при разработке надстроек, учитывайте, что конфигурация ESC включена по умолчанию в Windows Server.

- По умолчанию используется один из следующих браузеров: Internet Explorer 11 или последняя версия Microsoft Edge, Chrome, Firefox или Safari (Mac OS).
- Редактор HTML и JavaScript, например "Блокнот", [Visual Studio и Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs) или стороннее средство веб-разработки.

## <a name="client-requirements-os-x-desktop"></a>Требования к клиенту: настольный компьютер OS X

Outlook для Mac (входит в состав Office 365) поддерживает надстройки Outlook. При запуске надстроек Outlook в Outlook для Mac применяются те же требования, что и к Outlook для Mac: необходима операционная система OS X Yosemite версии 10.10 или более поздней. Так как Outlook для Mac использует WebKit в качестве обработчика макетов для преобразования страниц надстройки, то эта надстройка не зависит от браузеров.

Ниже приведены минимальные версии клиентов Office для Mac, которые поддерживают надстройки Office.

- Word для Mac версии 15.18 (160109);
- Excel для Mac версии 15.19 (160206);
- PowerPoint для Mac версии 15.24 (160614).

## <a name="client-requirements-browser-support-for-office-online-web-clients-and-sharepoint"></a>Требования к клиенту: поддержка браузеров для веб-клиентов Office Online и SharePoint

Любой браузер, поддерживающий ECMAScript 5.1, HTML5 и CSS3, например Internet Explorer 11 или последняя версия Microsoft Edge, Chrome, Firefox или Safari (Mac OS).


## <a name="client-requirements-non-windows-smartphone-and-tablet"></a>Требования к клиенту: смартфоны и планшеты под управлением операционных систем, отличных от Windows

Специально для приложения Outlook Web App, работающего в браузере на смартфонах и планшетах под управлением систем, отличных от Windows, для тестирования и запуска надстроек Outlook необходимо следующее программное обеспечение:


| Ведущее приложение | Устройство | Операционная система | Учетная запись Exchange | Мобильный браузер |
|:-----|:-----|:-----|:-----|:-----|
|Outlook для Android|Планшеты и смартфоны с Android|Android 4.4 KitKat или более поздней версии|Последнее обновление Office 365 для бизнеса или Exchange Online|Встроенное приложение для Android, браузер не применим|
|Outlook для iOS|Планшеты iPad, смартфоны iPhone|iOS 11 или более поздняя версия|Последнее обновление Office 365 для бизнеса или Exchange Online|Встроенное приложение для iOS, браузер не применим|
|Outlook Web App|iPhone 4, iPad 2, iPod Touch 4 или более поздние модели этих устройств|iOS 5 или более поздняя версия|Office 365, Exchange Online либо локальная среда Exchange Server 2013 или более поздней версии|Safari|

> [!NOTE]
> Встроенные приложения OWA для Android, OWA для iPad и OWA для iPhone [устарели](https://support.office.com/article/Microsoft-OWA-mobile-apps-are-being-retired-076ec122-4576-4900-bc26-937f84d25a4b) и больше не требуются и не применяются для тестирования надстроек Outlook.


## <a name="see-also"></a>См. также

- [Обзор платформы надстроек Office](../overview/office-add-ins.md)
- [Доступность ведущих приложений и платформ для надстроек Office](../overview/office-add-in-availability.md)
- [Веб-средства просмотра, используемые надстройками Office](web-viewers-used-by-office-web-add-ins.md)
