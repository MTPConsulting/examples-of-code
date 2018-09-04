class PostsController < ApplicationController
  before_action :authenticate_user!, except: [:index, :show]
  before_action :set_post, only: [:edit, :update, :destroy]
  before_action :check_permission_user, only: [:edit, :update, :destroy]

  def index
    conditions = ""
    # Si hay búsqueda obtengo el valor a filtrar
    q = params[:q]
    if q.present?
      conditions = "title ILIKE ?", "%#{q}%"
    end

    # Articulos generales
    @posts = Post.includes(:categories, :user)
    @posts = @posts.where(conditions).paginate(:page => params[:page])
    @posts = @posts.with_attached_image.order(created_at: :desc)

    # Queries para el layout
    @layout_queries = PostsLayoutQueries.new
  end

  def show
    @post = Post.includes(:categories, :comments, :user, comments: :user).with_attached_image
    @post = @post.where(id: params[:id], slug: params[:slug]).first

    if @post.present?
      # Queries para el layout
      @layout_queries = PostsLayoutQueries.new

      render 'posts/show'
    else
      raise ActionController::RoutingError.new("Not Found")
    end
  end

  def new
    @post = Post.new
  end

  def edit
  end

  def create
    @post = Post.new(post_params)

    respond_to do |format|
      @post.user = current_user
      if @post.save
        format.html { redirect_to posts_path, notice: 'Se creó el post correctamente.' }
      else
        format.html { render :new }
      end
    end
  end

  def update
    respond_to do |format|
      @post.user = current_user
      if @post.update(post_params)
        format.html { redirect_to post_path(@post.id, @post.slug), notice: 'Se editó el post correctamente.' }
      else
        format.html { render :edit }
      end
    end
  end

  def destroy
    @post.destroy
    redirect_to posts_path
  end

  private
    # Chequeo si tiene permisos
    def check_permission_user
      # El usuario logueado debe ser el mismo
      # que el usuario propietario del post
      if @post.user.username != current_user.username
        redirect_to posts_path
      end
    end

    def set_post
      @post = Post.find(params[:id])
    end

    def post_params
      params.require(:post).permit(:title, :body, :image, :category_ids => [])
    end
end
