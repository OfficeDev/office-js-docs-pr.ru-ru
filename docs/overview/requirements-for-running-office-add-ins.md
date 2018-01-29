
# <a name="requirements-for-running-office-add-ins"></a>Требования к запуску надстроек для Office


В этой статье описаны требования к программному обеспечению и устройствам для запуска надстроек Office.

> [!NOTE]
>  Если вы планируете [публиковать](../publish/publish.md) надстройку в Магазине Office, она должна соответствовать [политикам проверки Магазина Office](https://msdn.microsoft.com/ru-ru/library/jj220035.aspx), чтобы пройти проверку. Например, работать на всех платформах, поддерживающих определенные вами методы. Дополнительные сведения см. в [разделе 4.12](https://msdn.microsoft.com/ru-ru/library/jj220035.aspx#Anchor_3) и на [странице с информацией о доступности и ведущих приложениях для надстроек Office](https://dev.office.com/add-in-availability).

Общие сведения о том, в каких продуктах поддерживаются надстройки Office, см. на странице [Доступность ведущих приложений и платформ для Office](http://dev.office.com/add-in-availability).

## <a name="server-requirements"></a>Требования к серверу

Чтобы иметь возможность установить и запустить любую Надстройка Office, необходимо сначала развернуть файлы манифеста и веб-страниц для пользовательского интерфейса и кода надстройки в соответствующих папках на сервере.

Для всех типов надстроек (контентных надстроек, надстроек Outlook и надстроек области задач, а также команд надстроек) необходимо развертывать файлы веб-страниц на веб-сервере или в службе веб-хостинга, например [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md).


 >**Примечание.**   Когда вы создаете и отлаживаете надстройку в Visual Studio, эта система развертывает и запускает соответствующие файлы веб-страниц локально с помощью IIS Express, не требуя при этом использования дополнительного веб-сервера. 

Кроме того, требуется [каталог надстроек](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) в SharePoint, чтобы отправить XML-файл манифеста надстройки (контентной или области задач) в поддерживаемых ведущих приложениях Office — веб-приложениях Access, Word, Excel, PowerPoint и Project.

Чтобы тестировать и запускать надстройки Outlook, необходимо разместить учетную запись электронной почты Outlook в Exchange 2013 или более поздней версии, доступ к которой можно получить в Office 365, Exchange Online или в локально установленной версии. Пользователь или администратор устанавливают файлы манифестов надстроек Outlook на соответствующем сервере.

 >**Примечание.**   Учетные записи POP и IMAP в Outlook не поддерживают надстройки Office.




## <a name="client-requirements-windows-desktop-and-tablet"></a>Требования к клиенту: настольные компьютеры и планшеты под управлением Windows

Чтобы можно было разработать Надстройка Office для поддерживаемых классических клиентов Office или веб-клиентов, работающих на настольных компьютерах, ноутбуках или планшетах с ОС Windows, необходимо следующее программное обеспечение:


- Для настольных компьютеров под управлением 32- и 64-разрядных версий Windows, а также таких планшетов, как Surface Pro:
    - 32- или 64-разрядная версия Office 2013 или более поздняя версия в Windows 7 или более поздней версии.
    - Excel 2013, Outlook 2013, PowerPoint 2013, Project профессиональный 2013, Project 2013 с пакетом обновления 1 (SP1), Word 2013 или более поздняя версия клиента Office, если Надстройка Office тестируется или запускается специально для одного из этих клиентов Office. Клиенты Office для настольных ПК можно устанавливать локально или на клиентском компьютере с помощью технологии "нажми и работай".
    
        Если у вас не установлен Office 2013, но есть подписка на Office 365, вы можете скачать его из сети доставки содержимого по одной из следующих ссылок:
        
        - Office 2013 для бизнеса: [https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=O365BusinessRetail&platform=X86&language=en-us&version=O15GA&source=O15OLSO365](https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=O365BusinessRetail&platform=X86&language=en-us&version=O15GA&source=O15OLSO365) 
        - Office 2013 для дома: [https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=O365HomePremRetail&platform=X86&language=en-us&version=O15GA&source=O15OLSO365](https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=O365HomePremRetail&platform=X86&language=en-us&version=O15GA&source=O15OLSO365) 

- Браузер Internet Explorer 11 или более поздней версии должен быть установлен, но может не использоваться по умолчанию. Для поддержки надстроек Office клиент Office, выступающий в роли ведущего приложения, использует компоненты браузера, которые входят в состав Internet Explorer 11 или более поздней версии.
- По умолчанию используется один из следующих браузеров: Internet Explorer 11 или более поздней версии, последняя версия Microsoft Edge, Chrome, Firefox или Safari (Mac OS).
- Редактор HTML и JavaScript, например "Блокнот", [Visual Studio и Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs) или стороннее средство веб-разработки.

## <a name="client-requirements-os-x-desktop"></a>Требования к клиенту: настольный компьютер OS X

Outlook для Mac (входит в состав Office 365) поддерживает надстройки Outlook. При запуске надстроек Outlook в Outlook для Mac применяются те же требования, что и к Outlook для Mac: необходима операционная система OS X Yosemite версии 10.10 или более поздней. Так как Outlook для Mac использует WebKit в качестве обработчика макетов для преобразования страниц надстройки, то эта надстройка не зависит от браузеров.

Ниже приведены минимальные версии клиентов Office для Mac, которые поддерживают надстройки Office.

- Word для Mac версии 15.18 (160109); 
- Excel для Mac версии 15.19 (160206); 
- PowerPoint для Mac версии 15.24 (160614).

## <a name="client-requirements-browser-support-for-office-online-web-clients-and-sharepoint"></a>Требования к клиенту: поддержка браузеров для веб-клиентов Office Online и SharePoint

Любой браузер, поддерживающий ECMAScript 5.1, HTML5 и CSS3, например Internet Explorer 11 или более поздней версии, либо последняя версия Microsoft Edge, Chrome, Firefox или Safari (Mac OS).


## <a name="client-requirements-non-windows-smartphone-and-tablet"></a>Требования к клиенту: смартфоны и планшеты под управлением операционных систем, отличных от Windows

Специально для OWA для устройств и Outlook Web App, работающих в браузере на смартфонах и планшетах под управлением систем, отличных от Windows, для тестирования и запуска надстроек Outlook необходимо следующее программное обеспечение:


| Ведущее приложение | Устройство | Операционная система | Учетная запись Exchange | Мобильный браузер |
|:-----|:-----|:-----|:-----|:-----|
|OWA для Android|Смартфоны Android. Операционная система [Android OS](https://developer.android.com/guide/practices/screens_support.html) характеризует эти устройства как "небольшие" или "обычные".|Android 4.4 KitKat или более поздней версии|Последнее обновление Office 365 для бизнеса или Exchange Online|Встроенная надстройка для Android, браузер не применим|
|OWA для iPad|iPad 2 или более поздняя модель|iOS 6 или более поздняя версия|Последнее обновление Office 365 для бизнеса или Exchange Online|Встроенная надстройка для iOS, браузер не применим|
|OWA для iPhone|iPhone 4S или более поздняя модель|iOS 6 или более поздняя версия|Последнее обновление Office 365 для бизнеса или Exchange Online|Встроенная надстройка для iOS, браузер не применим|
|Outlook Web App|iPhone 4, iPad 2, iPod Touch 4 или более поздние модели этих устройств|iOS 5 или более поздняя версия|Office 365, Exchange Online либо локальная среда Exchange Server 2013 или более поздней версии|Safari|


## <a name="additional-resources"></a>Дополнительные ресурсы

- [Обзор платформы надстроек Office](../overview/office-add-ins.md)
- [Доступность ведущих приложений и платформ для надстроек Office](http://dev.office.com/add-in-availability)

