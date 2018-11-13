<%@ Page Title="" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true" CodeFile="Sheets.aspx.cs" Inherits="Sheets" %>

<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="Server">


    <div>
        <table>
            <tr>
                <td style="text-align: right;">Location: </td>
                <td>
                    <asp:TextBox ID="LocationBox" runat="server" /></td>
            </tr>
            <tr>
                <td style="text-align: right;">SKU: </td>
                <td>
                    <asp:TextBox ID="SKUbox" runat="server" /></td>
            </tr>

            <tr>
                <td style="text-align: right;">Quantity: </td>
                <td>
                    <asp:TextBox ID="QBox" runat="server" /></td>
            </tr>
            <tr>
                <td style="text-align: right;">
                    <asp:Button ID="AddBtn" runat="server" OnClick="AddBtn_Click" Text="Add" />
                </td>
                <td></td>
            </tr>
        </table>


    </div>


    <asp:Literal ID="MyMessage" runat="server" />



    <asp:DataGrid ID="MySheetGrid"
        AutoGenerateColumns="false"
        OnItemDataBound="MyDataGrid_ItemDataBound"
        OnDeleteCommand="MyDataGrid_Delete"
        OnUpdateCommand="MyDataGrid_Update"
        OnPageIndexChanged="MyData_Paged"
        OnSortCommand="MyDataGrid_Sort"
        AllowSorting="true"
        HeaderStyle-Font-Bold="true"
        HeaderStyle-HorizontalAlign="Center"
        AllowPaging="true"
        AlternatingItemStyle-BackColor="LightGray"
        PagerStyle-Mode="NumericPages"
        PagerStyle-Position="TopAndBottom"
        PageSize="15"
        runat="server">

        <Columns>
            <asp:BoundColumn HeaderText="Row" SortExpression="Row" DataField="Row" ItemStyle-Wrap="false" />
            <asp:BoundColumn HeaderText="Location" SortExpression="Location" DataField="Location" ItemStyle-Wrap="false" />
            <asp:BoundColumn HeaderText="SKU" SortExpression="SKU" DataField="SKU" ItemStyle-Wrap="false" />
            <asp:TemplateColumn HeaderText="QTY" SortExpression="QTY">
                <ItemTemplate>
                    <asp:TextBox Width="100px" runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "QTY").ToString()%>' />

                </ItemTemplate>
            </asp:TemplateColumn>
            <asp:ButtonColumn ButtonType="PushButton" Text="Update" CommandName="Update" />
            <asp:ButtonColumn ButtonType="PushButton" Text="Delete" CommandName="Delete" />

        </Columns>

    </asp:DataGrid>


</asp:Content>

