<script lang="ts">

    import * as XLSX from 'xlsx';
    import CardFront from '../components/CardFront.svelte';
    import CardBack from '../components/CardBack.svelte';

    let tableData: any = [];
    let headers: any = [];
    let selectedRows: number[] = [];

    const handleFileUpload = async (event: any) => {
        const file = event.target.files[0];

        if (!file || !file.name.endsWith('.xlsx')) {
            alert('Please upload a valid .xlsx file.');
            return;
        }

        const reader = new FileReader();

        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // Assuming the first sheet is the one we want
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];

            // Convert sheet data to JSON
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            if (jsonData.length) {
                headers = jsonData[0]; // First row as headers
                tableData = jsonData.slice(1); // Remaining rows as data
                selectedRows = []; // Reset selected rows when new data is loaded
            }
        };

        reader.readAsArrayBuffer(file);
    };

    // Function to toggle row selection
    function toggleRowSelection(rowIndex: number) {
        if (selectedRows.includes(rowIndex)) {
            selectedRows = selectedRows.filter(index => index !== rowIndex);
        } else {
            selectedRows = [...selectedRows, rowIndex];
        }
    }

    // Function to toggle select/deselect all rows
    function toggleSelectAll() {
        if (selectedRows.length === tableData.length) {
            // Deselect all rows
            selectedRows = [];
        } else {
            // Select all rows
            selectedRows = tableData.map((_, index) => index);
        }
    }

    // Reactive statement to determine if all rows are selected
    $: allRowsSelected = selectedRows.length === tableData.length && tableData.length > 0;

    function getDataForRow(rowIndex: number) {
        const row = tableData[rowIndex];
        const data: any = {};
        headers.forEach((header: string, i: number) => {
            data[header] = row[i];
        });
        return data;
    }

    function printCards() {
        // Prepare the content to print
        const printableElement = document.createElement('div');

        selectedRows.forEach(rowIndex => {
            const data = getDataForRow(rowIndex);

            // Create a container div for each card set
            const cardSet = document.createElement('div');

            // Render CardFront
            const cardFront = document.createElement('div');
            new CardFront({
                target: cardFront,
                props: { data }
            });
            cardSet.appendChild(cardFront);

            // Page break after front side
            const pageBreakFront = document.createElement('div');
            pageBreakFront.classList.add('page-break');
            cardSet.appendChild(pageBreakFront);

            // Render CardBack
            const cardBack = document.createElement('div');
            new CardBack({
                target: cardBack,
                props: { data }
            });
            cardSet.appendChild(cardBack);

            // Page break after back side
            const pageBreakBack = document.createElement('div');
            pageBreakBack.classList.add('page-break');
            cardSet.appendChild(pageBreakBack);

            printableElement.appendChild(cardSet);
        });

        // Create an iframe
        const iframe = document.createElement('iframe');
        iframe.style.display = 'none';
        document.body.appendChild(iframe);

        // Write the content to the iframe
        const doc = iframe.contentWindow.document;
        doc.open();
        doc.write(`
            <html>
                <head>
                    <title>Print Cards</title>
                    <style>
                        ${Array.from(document.querySelectorAll('style'))
                            .map(style => style.innerHTML)
                            .join('\n')}

                        @media print {
                            .page-break {
                                page-break-after: always;
                            }
                        }
                    </style>
                </head>
                <body>
                    ${printableElement.innerHTML}
                </body>
            </html>
        `);
        doc.close();

        // Wait for the content to load
        iframe.onload = function() {
            iframe.contentWindow.focus();
            iframe.contentWindow.print();
            document.body.removeChild(iframe);
        };
    }
</script>

<div class="w-full h-full flex">
    <div class="flex-1 border-r max-w-[50%] overflow-y-scroll">
        <div class="border-b p-2 w-full">
            <div class="w-[300px]">
                <input
                    id="file"
                    type="file"
                    accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    on:change={handleFileUpload}
                />
            </div>
        </div>
        <div class="p-2">
            {#if selectedRows.length}
                <button color="orange" on:click={printCards}>
                    Print ({selectedRows.length})
                </button>
            {/if}
        </div>
        {#if tableData.length}
            <table border="1" style="margin-top: 20px; width: 100%; text-align: left;">
                <thead>
                    <tr>
                        <th>
                            <input
                                type="checkbox"
                                checked={allRowsSelected}
                                on:change={toggleSelectAll}
                            />
                        </th>
                        {#each headers as header, i (i)}
                            <th>{header}</th>
                        {/each}
                    </tr>
                </thead>
                <tbody>
                    {#each tableData as row, rowIndex (rowIndex)}
                        <tr class:selected={selectedRows.includes(rowIndex)}>
                            <td>
                                <input
                                    type="checkbox"
                                    checked={selectedRows.includes(rowIndex)}
                                    on:change={() => toggleRowSelection(rowIndex)}
                                />
                            </td>
                            {#each row as cell}
                                <td>{cell}</td>
                            {/each}
                        </tr>
                    {/each}
                </tbody>
            </table>
        {/if}
    </div>
    <div class="flex-1 bg-slate-50 flex items-center justify-center relative">
        <!-- You can display a preview here if needed -->
    </div>
</div>

<style>
    th, td {
        min-width: 50px;
        border: 1px solid #e1e1e1;
        padding: 5px;
        font-size: 13px;
    }

    /* Highlight selected rows */
    tr.selected {
        background-color: #e0f7fa; /* Light cyan background for selected rows */
    }

    .card td {
        padding:1px !important;
        text-align: left;
    }

    .page-break {
        page-break-after: always;
    }

    /* Styles for CardFront and CardBack */
    .card-container {
        width: 5.70375in;
        height: 3.59125in;
        position: relative;
        background-size: 100%;
        background-repeat: no-repeat;
    }

    /* Additional styles if needed */
</style>
