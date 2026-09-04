import nodemailer from "nodemailer";

function required(name: string) {
  const value = process.env[name]?.trim();
  if (!value) throw new Error(`${name} no está configurada en el servidor.`);
  return value;
}

export async function sendMinutesEmail(input: { recipients: string[]; subject: string; text: string }) {
  const host = required("SMTP_HOST");
  const port = Number(process.env.SMTP_PORT || 587);
  const user = required("SMTP_USER");
  const password = required("SMTP_PASSWORD");
  const from = (process.env.SMTP_FROM || user).trim();
  const transporter = nodemailer.createTransport({ host, port, secure: process.env.SMTP_SECURE === "true", auth: { user, pass: password } });
  await transporter.sendMail({ from, to: input.recipients.join(","), subject: input.subject, text: input.text });
  return { sent: true, from, recipients: input.recipients };
}
