module.exports = {
  provider: "microsoft-graph",
  tenantId: "3046c5ed-044f-4ada-9ad4-d4e085cb4cb6",
  clientId: "bc1465b4-e922-4fad-94ac-a13861d7ad36",
  clientSecret: "bd7c3d11-0723-4e8e-b00f-303398186e99",
  from: "artwork3@giftwrap.co.za",
  to: "admin3@giftwrap.co.za",
  artworkTo: process.env.MAIL_ARTWORK_TO || "artwork3@giftwrap.co.za",
};
