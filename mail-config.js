module.exports = {
  disabled: false,
  provider: "microsoft-graph",
  tenantId: process.env.MAIL_TENANT_ID || "3046c5ed-044f-4ada-9ad4-d4e085cb4cb6",
  clientId: process.env.MAIL_CLIENT_ID || "bc1465b4-e922-4fad-94ac-a13861d7ad36",
  clientSecret: process.env.MAIL_CLIENT_SECRET || "vXR8Q~s-eXyhb~LYg9DW0i5nAfXavhvhZcaPJau4",
  from: "artwork3@giftwrap.co.za",
  to: "admin3@giftwrap.co.za, promo10@giftwrap.co.za",
  artworkTo: process.env.MAIL_ARTWORK_TO || "artwork3@giftwrap.co.za",
};
