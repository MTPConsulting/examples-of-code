module PaginateHelper
  class LinkRenderer < WillPaginate::ActionView::LinkRenderer
    def container_attributes
      {class: "pagination"}
    end

    def page_number(page)
      "<li class=\"page-item\">" + link(page, page, :rel => rel_value(page), :class => "page-link") + "</li>"
    end

    def previous_page
      num = @collection.current_page > 1 && @collection.current_page - 1
      previous_or_next_page(num, '<span aria-hidden="true">&laquo;</span>', "page-link")
    end

    def next_page
      num = @collection.current_page < total_pages && @collection.current_page + 1
      previous_or_next_page(num, '<span aria-hidden="true">&raquo;</span>', "page-link")
    end

    def previous_or_next_page(page, text, classname)
      if page
        "<li class=\"page-item\">" + link(text, page, :class => classname) + "</li>"
      else
        "<li class=\"page-item disabled\">" + link(text, page, :class => classname) + "</li>"
      end
    end
  end
end

