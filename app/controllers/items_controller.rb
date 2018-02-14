def item_params
    params.require(:item).permit(:name, :description :picture) # Add :picture as a permitted paramter
end
