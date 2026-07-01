export interface EmailAddress {
  name?: string;
  address: string;
}

export interface PhysicalAddress {
  street?: string;
  city?: string;
  state?: string;
  countryOrRegion?: string;
  postalCode?: string;
}

/** Subset of the Microsoft Graph `contact` resource we read and write. */
export interface Contact {
  id: string;
  displayName?: string;
  givenName?: string;
  surname?: string;
  nickName?: string;
  emailAddresses?: EmailAddress[];
  mobilePhone?: string | null;
  homePhones?: string[];
  businessPhones?: string[];
  homeAddress?: PhysicalAddress;
  businessAddress?: PhysicalAddress;
  otherAddress?: PhysicalAddress;
  companyName?: string;
  jobTitle?: string;
  department?: string;
  birthday?: string | null;
  personalNotes?: string;
  parentFolderId?: string;
  lastModifiedDateTime?: string;
}
