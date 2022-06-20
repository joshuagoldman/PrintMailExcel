namespace MailInboxApi.DataAccess
{
    public readonly record struct MailClient(string EmailAddress, string Password);
}
