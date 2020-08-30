$('#upload').click(function (e) {
    e.preventDefault(); // stop the standard form submission

    var fd = new FormData();
    var fileUpload = $('[name=files]').get(0);
    var files = fileUpload.files;

    if (files.length > 0)
        fd.append("file", files[0]);

    $.ajax({
        url: '/Home/UploadFile',
        method: 'POST',
        data: fd,
        contentType: false,
        processData: false,
        success: function (data) {
            console.log(data);
            if (Array.isArray(data)) {
                let ddSheets = document.querySelector('.row:nth-child(2) select');
                $(ddSheets).empty();
                $(ddSheets).append('<option>--Select Sheet--</option>');
                for (let sheet in data)
                    $(ddSheets).append('<option>' + data[sheet] + '</option>');
            } else
                alert("Unable to fetch sheets. Please verify file and retry.");
        },
        error: function (xhr, error, status) {
            alert(status + "-" + error);
        }
    });
});

function getColumns(sheetname, isColumnNames) {
    $.ajax({
        url: '/Home/GetColumns',
        method: 'GET',
        data: { sheetName: sheetname, firstrow: isColumnNames },
        success: function (data) {
            let table = document.querySelector('.custab');
            $(table).children('tr').remove();
            for (let column in data) {
                $(table).append(`
                                    <tr>
                                        <td>
                                            <input type='checkbox' />
                                        </td>
                                        <td>`+ data[column] + `</td>
                                        <td>
                                            <select class="form-control">
                                                <option>String</option>
                                                <option>DateTime</option>
                                            </select>
                                        </td>
                                        <td>
                                            <input type='text' class='form-control' placeholder='Special characters to replace' />
                                        </td>
                                    </tr>`);
            }
        },
        error: function (xhr, error, status) {
            alert(status + "-" + error);
        }
    });
}

$('.row:nth-child(2) input[type=checkbox]').on('change', function () {
    if ($(this).siblings('select').children('option:selected').index() > 0)
        getColumns($(this).siblings('select').children('option:selected').text(), $(this).is(':checked'));
});

$('.row:nth-child(2) select').on('change', function () {
    if ($(this).children('option:selected').index() > 0)
        getColumns($(this).children('option:selected').text(), $('.row:nth-child(2) input[type=checkbox]').is(':checked'));
});

$('body').on('change', '.row:nth-child(3) select', function () {
    debugger;
    let index = $(this).index();
    if ($(this).children('option:selected').index() == 1)
        $(this).closest('tr').find('input').attr('placeholder', 'DateTime format in Excel - Ex: MM/dd/yyyy HH:mm:ss');
    else
        $(this).closest('tr').find('input').attr('placeholder', 'Special characters to replace');
});

$('.row:last button').on('click', function () {
    let btns = $(this).closest('.basebtn').find('button');
    let btnIndex = btns.index($(this));
    let dbRow = document.querySelectorAll('.row:nth-child(4) input');
    let data = {
        SheetName: $('.row:nth-child(2) select').children('option:selected').text(),
        Firstrow: $('.row:nth-child(2) input[type=checkbox]').is(':checked'),
        Rows: [],
        DBDetails: {
            TableName: dbRow[0].innerText,
            InstanceName: dbRow[1].innerText,
            InitialCatalog: dbRow[2].innerText,
            Password: dbRow[3].innerText
        }
    };

    let rows = $('table > tr');

    for (var i = 0; i < rows.length; i++) {
        let row = {};
        let curRowCells = rows[i].querySelectorAll('td');
        if (!curRowCells[0].querySelector('input').checked)
            continue;
        row.ColumnName = curRowCells[1].innerText;
        row.DataType = $(curRowCells[2].querySelector('select')).children('option:selected').text();
        row.ExtraInfo = curRowCells[3].querySelector('input').value;
        data.Rows.push(row);
    }

    switch (btnIndex) {
        case 0:
            break;
        case 1:
            $.ajax({
                url: '/Home/GetScripts',
                method: 'POST',
                data: JSON.stringify(data),
                contentType: "application/json",
                success: function (data) {
                    if (data)
                        window.open("/Home/Download?fileName=" + data, "_blank");
                },
                error: function (xhr, error, status) {
                    alert(status + "-" + error);
                }
            });
            break;
        default:
    }
});