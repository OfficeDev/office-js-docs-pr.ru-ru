# <a name="method-element"></a>Элемент Method

Указывает отдельный метод из API JavaScript для Office, необходимый для активации надстройки Office.

**Тип надстройки:** содержимое, область задач.

## <a name="syntax"></a>Синтаксис

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>Содержится в

[Методы](methods.md)

## <a name="attributes"></a>Атрибуты

|**Атрибут**|**Тип**|**Обязательный**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя|string|обязательный|Указывает имя необходимого метода, соответствующее его родительскому объекту. Например, чтобы задать метод **getSelectedDataAsync**, необходимо указать `"Document.getSelectedDataAsync"`.|

## <a name="remarks"></a>Замечания

Элементы  **Methods** и **Method** не поддерживаются надстройками почты. Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

> [!IMPORTANT] 
> Минимальную версию невозможно указать для отдельных методов. Чтобы убедиться, что метод доступен в среде выполнения, при вызове этого метода в сценарии надстройки следует также использовать оператор **if**. Дополнительные сведения о том, как это сделать, см. в статье [Общие сведения об API JavaScript для Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).

