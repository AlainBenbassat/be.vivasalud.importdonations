<h3>Import Log</h3>
<div class="crm-block crm-content-block">
    <table class="crm-info-panel">
        <tr class="columnheader">
            <th>ID</th>
            <th>Worksheet</th>
            <th>Cell</th>
            <th>Comment Type</th>
            <th>Comment</th>
        </tr>
        {foreach from=$log item=row}
            <tr class="{cycle values="odd-row,even-row"}">
                <td>{$row.id}</td>
                <td>{$row.worksheet}</td>
                <td>{$row.cell}</td>
                <td>{$row.comment_type}</td>
                <td>{$row.comment}</td>
            </tr>
        {/foreach}
    </table>
</div>
