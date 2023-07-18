/** a link contains a href URL and HTTP method */
type HrefMethod = {
  href: string,
  method: string
};

/** a self link using rel self */
export type SelfLink = {
  self: HrefMethod,
};

/** a result of type T which has a links containing a self-link */
type LinkedResult<T> = {
  links: SelfLink,
  result: T,
};

/** a response envelope always has an isOk and HTTP status */
type Envelope = {
  isOk: boolean,
  status: number,
};

/** an envelope for a successful response */
export type SuccessEnvelope<T> = Envelope & LinkedResult<T> & {
  isOk: true,
};

/** an envelope for a failed response */
export type ErrorEnvelope = Envelope & {
  isOk: false,
  errors: { message: string, options?: { [key:string]: string } }[],
};
