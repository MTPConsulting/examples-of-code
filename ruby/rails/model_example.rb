require 'minidown'

class Post < ApplicationRecord
  before_save :set_slug

  has_and_belongs_to_many :categories
  belongs_to :user
  has_one_attached :image
  has_many :comments, dependent: :destroy

  validates :title, presence: true
  validates :body, presence: true
  validates :user, presence: true
  validates :image, presence: true
  validates :category_ids, presence: true
  validate :image_validation

  # Pagination
  self.per_page = 10

  # Para parsear el markdown completo
  def body_markdown
    Minidown.render(self.body).html_safe
  end

  private
    def image_validation
      if image.attached?
        if image.blob.byte_size > 1000000
          image.purge
          errors[:base] << 'Too big'
        elsif !image.blob.content_type.starts_with?('image/')
          image.purge
          errors[:base] << 'Wrong format'
        end
      end
    end

    def set_slug
      self.slug = self.title.parameterize
    end
end
