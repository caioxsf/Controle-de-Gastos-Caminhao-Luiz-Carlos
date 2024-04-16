    document.addEventListener('DOMContentLoaded', function() {
        var data = JSON.parse(localStorage.getItem('tabelaPronta')) || [];
        var tabelaPronta = document.getElementById('tabelaPronta');

        data.forEach(function(colunaData) {
            var row = tabelaPronta.insertRow(-1);
            colunaData.forEach(function(cellData) {
                var cell = row.insertCell();
                cell.textContent = cellData;
            });

            var linhaAcao = row.insertCell();
            var botaoDelete = document.createElement('button');
            botaoDelete.textContent = 'Apagar';
            botaoDelete.onclick = function() {
                tabelaPronta.deleteRow(row.rowIndex);
                salvarTabela();
            };
            // outro btn do apagar (nao mexer oreia curioso)
            linhaAcao.appendChild(botaoDelete);
        });
        });

    function addColuna() {
        var dataInput = document.getElementById('data').value;
        var data = moment(dataInput, 'YYYY-MM-DD').format('DD/MM/YYYY');
        var origem = document.getElementById('origem').value;
        var destino = document.getElementById('destino').value;
        var frete = document.getElementById('frete').value;
        var material = document.getElementById('material').value;
        var peso = document.getElementById('peso').value;
        var motorista = document.getElementById('motorista').value;
        var transportadora = document.getElementById('transportadora').value;
        // var adiantamento = document.getElementById('adiantamento').value;

        var valor = frete * peso;
        var comissao = valor * 8 / 100;
        var valor_formatado = valor.toLocaleString('pt-BR', {style: 'currency', currency: 'BRL'});
        var comissao_formatado = comissao.toLocaleString('pt-BR', {style: 'currency', currency: 'BRL'});

        // valor_total=0;
        // valor_total = comissao + valor_total;
        // var adiantamento_valor = adiantamento - valor_total;

        var tabelaPronta = document.getElementById('tabelaPronta');
        var row = tabelaPronta.insertRow(-1);
        var cell1 = row.insertCell();
        var cell2 = row.insertCell();
        var cell3 = row.insertCell();
        var cell4 = row.insertCell();
        var cell5 = row.insertCell();
        var cell6 = row.insertCell();
        var cell7 = row.insertCell();
        var cell8 = row.insertCell();
        var cell9 = row.insertCell();
        var cell10 = row.insertCell();
        // var cell11 = row.insertCell();

        cell1.textContent = data;
        cell2.textContent = origem;
        cell3.textContent = destino;
        cell4.textContent = frete;
        cell5.textContent = material;
        cell6.textContent = peso;
        cell7.textContent = motorista;
        cell8.textContent = valor_formatado;
        cell9.textContent = comissao_formatado;
        cell10.textContent = transportadora;
        // cell11.textContent = adiantamento_valor;

        // funcao do botao para apagar linha por linha (retirei pq eu quis)

        var linhaAcao = row.insertCell();
        var botaoDelete = document.createElement('button');
        botaoDelete.textContent = 'Apagar';
        botaoDelete.onclick = function() {
            tabelaPronta.deleteRow(row.rowIndex);
            salvarTabela();
        };
        linhaAcao.appendChild(botaoDelete);

        salvarTabela();

        botaoDelete.className = 'botao-apagar';
    }

    function salvarTabela() {
        var tabelaPronta = document.getElementById('tabelaPronta');
        var data = [];
        for (var i = 0; i < tabelaPronta.rows.length; i++) {
            var colunaData = [];
            for (var j = 0; j < tabelaPronta.rows[i].cells.length - 1; j++) {
                colunaData.push(tabelaPronta.rows[i].cells[j].textContent);
            }
            data.push(colunaData);
        }
        localStorage.setItem('tabelaPronta', JSON.stringify(data));
    }

    function clearTable() {
        var tabelaPronta = document.getElementById('tabelaPronta');
        while (tabelaPronta.firstChild) {
            tabelaPronta.removeChild(tabelaPronta.firstChild);
        }
        localStorage.removeItem('tabelaPronta');
    }

    function exportToExcel() {
        var table = document.getElementById('tabelaForm');
        var rows = table.rows;
        var data = [];
        for (var i = 0; i < rows.length; i++) {
            var colunaData = [];
            var cells = rows[i].cells;
            for (var j = 0; j < cells.length - 1; j++) {
                colunaData.push(cells[j].textContent);
            }
            data.push(colunaData);
        }

        var wb = XLSX.utils.book_new();
        var ws = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, ws, 'Tabela');

        var wbout = XLSX.write(wb, {bookType:'xlsx',  type: 'binary'});
        function s2ab(s) {
            var buf = new ArrayBuffer(s.length);
            var view = new Uint8Array(buf);
            for (var i=0; i<s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
            return buf;
        }

        var blob = new Blob([s2ab(wbout)],{type:"application/octet-stream"});
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = 'planilha.xlsx';
        a.click();
        setTimeout(function() { URL.revokeObjectURL(url); }, 0);
    }
