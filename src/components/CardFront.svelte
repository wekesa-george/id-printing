<!-- CardFront.svelte -->
<script>
    export let data;
    import Img from "$lib/images/card.png";
</script>

<div class="w-[5.70375in] h-[3.59125in] bg-white border rounded-lg shadow-md relative"
     style="background-image:url({Img}); background-size:100%;background-repeat:no-repeat;">
    <div class="border-l-[7px] border-l-red-600 w-[250px]  absolute top-[120px] left-3">
        <table class="card">
            <tbody>
                <tr>
                    <td class="bg-green-600 text-white font-medium text-[14px] ">
                        Name :
                    </td>
                    <td class="">
                        {data['Name']}
                    </td>
                </tr>
                <tr>
                    <td class="bg-green-600 text-white font-medium text-[14px] ">
                        Designation :
                    </td>
                    <td class="">
                        {data['Designation']}
                    </td>
                </tr>
                <tr>
                    <td class="bg-green-600 text-white font-medium text-[14px] ">
                        Employee No. :
                    </td>
                    <td class="card">
                        {data['Employee No.']}
                    </td>
                </tr>
                <tr>
                    <td class="bg-green-600 text-white font-medium text-[14px] ">
                        Issue Date :
                    </td>
                    <td class="">
                        {data['Issue Date']}
                    </td>
                </tr>
                <tr>
                    <td class="bg-green-600 text-white font-medium text-[14px] ">
                        Expiry :
                    </td>
                    <td class="">
                        {data['Expiry']}
                    </td>
                </tr>
            </tbody>
        </table>
    </div>
</div>
