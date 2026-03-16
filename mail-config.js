module.exports = {
  provider: "microsoft-graph",
  tenantId: process.env.MAIL_TENANT_ID || "3046c5ed-044f-4ada-9ad4-d4e085cb4cb6",
  clientId: process.env.MAIL_CLIENT_ID || "bc1465b4-e922-4fad-94ac-a13861d7ad36",
  clientSecret: process.env.MAIL_CLIENT_SECRET || "bd7c3d11-0723-4e8e-b00f-303398186e99",
  from: "artwork3@giftwrap.co.za",
  to: "admin3@giftwrap.co.za",
};
