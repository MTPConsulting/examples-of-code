class ExampleMailer < ApplicationMailer

  def account_activation(user)
    @user = user
    mail to: user.email, subject: "Activation account"
  end

  def password_reset(user)
    @user = user
    mail to: user.email, subject: "Password reset"
  end

  def contact_us(user, data, subject)
    @body = data
    @user = user
    mail to: "test@test", subject: subject
  end
end
