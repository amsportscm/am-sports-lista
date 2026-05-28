export class Readable {
  pipe() { return this; }
  on() { return this; }
  resume() { return this; }
}
export class Writable {
  write() { return true; }
  end() { return this; }
  on() { return this; }
}
export class Transform extends Readable {}
const stream = { Readable, Writable, Transform };
export default stream;
