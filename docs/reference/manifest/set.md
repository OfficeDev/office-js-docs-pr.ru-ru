# <a name="set-element"></a>Элемент Set

Указывает набор требований из API JavaScript для Office, необходимый для активации надстройки Office.

**Тип надстройки:** содержимое, область задач, почта

## <a name="syntax"></a>Синтаксис

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>Содержится в

[Наборы](sets.md)

## <a name="attributes"></a>Атрибуты

|**Атрибут**|**Тип**|**Обязательный**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя|string|обязательный|Имя [набора требований](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).|
|MinVersion|string|необязательный|Указывает минимальную версию набора API, необходимую надстройке. Переопределяет значение **DefaultMinVersion**, если оно указано в родительском элементе [Sets](sets.md).|

## <a name="remarks"></a>Замечания

Дополнительные сведения о наборах требований см. в статье [версии и наборы требований  Office](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Дополнительные сведения об атрибуте **MinVersion** элемента **Set** и атрибуте **DefaultMinVersion** элемента **Sets** см. в статье [Указание элемента Requirements в манифесте](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).

> [!IMPORTANT] 
> Для надстроек почты существует только один `"Mailbox"` набор требований. Этот набор обязательных элементов содержит целое подмножество API, поддерживаемых в надстройках почты для Outlook, и необходимо указать `"Mailbox"` набор обязательных требований в манифесте надстройки почты (это необходимо, как в случае содержимого, так и надстройки области задач). Кроме того, невозможно объявить поддержку для отдельных методов в надстройках почты.
