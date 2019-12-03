<%@ Page Language="C#" AutoEventWireup="true" CodeFile="lavScheduler.aspx.cs" Inherits="lavScheduler" %>

<%@ Register assembly="DevExpress.Web.ASPxScheduler.v15.2, Version=15.2.9.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxScheduler" tagprefix="dxwschs" %>
<%@ Register assembly="DevExpress.XtraScheduler.v15.2.Core, Version=15.2.9.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.XtraScheduler" tagprefix="cc1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body style="font-family: Cambria;">
    <form id="form1" runat="server">
    <div style="text-align:center;">
        <dxwschs:ASPxScheduler ID="ASPxScheduler1" runat="server" AppointmentDataSourceID="SqlDataSource1" ClientIDMode="AutoID" ResourceDataSourceID="SqlDataSource2" Start="2017-05-04" 
            ActiveViewType="Week" Theme="Moderno" >
            <Storage>
                <Appointments AutoRetrieveId="True">
                    <Mappings AllDay="AllDay" AppointmentId="UniqueID" Description="Description" End="EndDate" Label="Label" Location="Location" RecurrenceInfo="RecurrenceInfo" ReminderInfo="ReminderInfo" ResourceId="ResourceID" Start="StartDate" Status="Status" Subject="Subject" Type="Type" />
                    <CustomFieldMappings>
                        <dxwschs:ASPxAppointmentCustomFieldMapping Member="ResourceIDs" Name="ResourceIDs" />
                        <dxwschs:ASPxAppointmentCustomFieldMapping Member="CustomField1" Name="CustomField1" />
                    </CustomFieldMappings>
                </Appointments>
                <Resources>
                    <Mappings Caption="ResourceName" Color="Color" Image="Image" ResourceId="ResourceID" />
                </Resources>
            </Storage>
            <views>
                <DayView DisplayName="Calendario giornaliero" MenuCaption="Vista giornaliera" ShortDisplayName="Giorno"><TimeRulers>
                <cc1:TimeRuler></cc1:TimeRuler>
                </TimeRulers>
                    <DayViewStyles ScrollAreaHeight="600px">
                    </DayViewStyles>
                </DayView>
                <WorkWeekView DisplayName="Calendario lavorativi" MenuCaption="Vista lavorativi" ShortDisplayName="Lavorativi">
                    <TimeRulers>
                        <cc1:TimeRuler></cc1:TimeRuler>
                    </TimeRulers>
                </WorkWeekView>
                <WeekView DisplayName="Calendario settimana/giorni" MenuCaption="Vista settimana/giorni" ShortDisplayName="Settimana/Giorni">
                </WeekView>
                <MonthView DisplayName="Calendario mensile" MenuCaption="Vista mensile" ShortDisplayName="Mese">
                </MonthView>
                <TimelineView DisplayName="Calendiario Timeline" MenuCaption="Vista Timeline">
                </TimelineView>
                <fullweekview enabled="true" DisplayName="Calendario settimana/ore" MenuCaption="Vista settimana/ore" ShortDisplayName="Settimana/Ore">
                    <TimeRulers>
                        <cc1:TimeRuler></cc1:TimeRuler>
                    </TimeRulers>
                </fullweekview>
            </views>
        </dxwschs:ASPxScheduler>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>" SelectCommand="SELECT * FROM [Resources]"></asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>" DeleteCommand="DELETE FROM [Appointments] WHERE [UniqueID] = @UniqueID" InsertCommand="INSERT INTO [Appointments] ([Type], [StartDate], [EndDate], [AllDay], [Subject], [Location], [Description], [Status], [Label], [ResourceID], [ResourceIDs], [ReminderInfo], [RecurrenceInfo], [CustomField1]) VALUES (@Type, @StartDate, @EndDate, @AllDay, @Subject, @Location, @Description, @Status, @Label, @ResourceID, @ResourceIDs, @ReminderInfo, @RecurrenceInfo, @CustomField1)" SelectCommand="SELECT * FROM [Appointments]" UpdateCommand="UPDATE [Appointments] SET [Type] = @Type, [StartDate] = @StartDate, [EndDate] = @EndDate, [AllDay] = @AllDay, [Subject] = @Subject, [Location] = @Location, [Description] = @Description, [Status] = @Status, [Label] = @Label, [ResourceID] = @ResourceID, [ResourceIDs] = @ResourceIDs, [ReminderInfo] = @ReminderInfo, [RecurrenceInfo] = @RecurrenceInfo, [CustomField1] = @CustomField1 WHERE [UniqueID] = @UniqueID">
            <DeleteParameters>
                <asp:Parameter Name="UniqueID" Type="Int32" />
            </DeleteParameters>
            <InsertParameters>
                <asp:Parameter Name="Type" Type="Int32" />
                <asp:Parameter Name="StartDate" Type="DateTime" />
                <asp:Parameter Name="EndDate" Type="DateTime" />
                <asp:Parameter Name="AllDay" Type="Boolean" />
                <asp:Parameter Name="Subject" Type="String" />
                <asp:Parameter Name="Location" Type="String" />
                <asp:Parameter Name="Description" Type="String" />
                <asp:Parameter Name="Status" Type="Int32" />
                <asp:Parameter Name="Label" Type="Int32" />
                <asp:Parameter Name="ResourceID" Type="Int32" />
                <asp:Parameter Name="ResourceIDs" Type="String" />
                <asp:Parameter Name="ReminderInfo" Type="String" />
                <asp:Parameter Name="RecurrenceInfo" Type="String" />
                <asp:Parameter Name="CustomField1" Type="String" />
            </InsertParameters>
            <UpdateParameters>
                <asp:Parameter Name="Type" Type="Int32" />
                <asp:Parameter Name="StartDate" Type="DateTime" />
                <asp:Parameter Name="EndDate" Type="DateTime" />
                <asp:Parameter Name="AllDay" Type="Boolean" />
                <asp:Parameter Name="Subject" Type="String" />
                <asp:Parameter Name="Location" Type="String" />
                <asp:Parameter Name="Description" Type="String" />
                <asp:Parameter Name="Status" Type="Int32" />
                <asp:Parameter Name="Label" Type="Int32" />
                <asp:Parameter Name="ResourceID" Type="Int32" />
                <asp:Parameter Name="ResourceIDs" Type="String" />
                <asp:Parameter Name="ReminderInfo" Type="String" />
                <asp:Parameter Name="RecurrenceInfo" Type="String" />
                <asp:Parameter Name="CustomField1" Type="String" />
                <asp:Parameter Name="UniqueID" Type="Int32" />
            </UpdateParameters>
        </asp:SqlDataSource>
    </div>
    </form>
</body>
</html>
